"""
Reyart Lead Sistemi — Streamlit Arayüz
Çalıştırma: streamlit run ui.py
"""

import json
import subprocess
import sys
import re
from pathlib import Path
from datetime import datetime, date
from io import BytesIO

import pandas as pd
import streamlit as st

# ── Ayarlar ──────────────────────────────────────────────────────────────────

OUTPUT_DIR = Path("output")
FAIRS_FILE = OUTPUT_DIR / "fairs.json"
COMPANIES_FILE = OUTPUT_DIR / "companies.json"
LEADS_FILE = OUTPUT_DIR / "leads.xlsx"

st.set_page_config(
    page_title="Reyart Fuar Lead Sistemi",
    page_icon="🏭",
    layout="wide",
)


# ── Veri yükleme ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=300)
def load_fairs() -> pd.DataFrame:
    if not FAIRS_FILE.exists():
        return pd.DataFrame()
    data = json.loads(FAIRS_FILE.read_text(encoding="utf-8"))
    df = pd.DataFrame(data)
    if "start_date" in df.columns:
        df["start_dt"] = pd.to_datetime(df["start_date"], errors="coerce")
        df["end_dt"] = pd.to_datetime(df["end_date"], errors="coerce")
    return df


@st.cache_data(ttl=60)
def load_leads() -> pd.DataFrame:
    if not COMPANIES_FILE.exists():
        return pd.DataFrame()
    data = json.loads(COMPANIES_FILE.read_text(encoding="utf-8"))
    return pd.DataFrame(data)


# ── Yardımcılar ──────────────────────────────────────────────────────────────

def refresh_calendar():
    with st.spinner("TOBB Excel'i indiriyor ve fuar takvimi güncelleniyor..."):
        proc = subprocess.run(
            [sys.executable, "fair_calendar.py"],
            capture_output=True, text=True, timeout=300,
        )
        st.cache_data.clear()
        if proc.returncode == 0:
            st.success("Fuar takvimi güncellendi!")
        else:
            st.error(f"Hata: {proc.stderr[-500:]}")


def backup_companies():
    """companies.json'u zaman damgalı yedeğe kopyala — veri kaybını önler."""
    if not COMPANIES_FILE.exists():
        return None
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = OUTPUT_DIR / f"companies_backup_{ts}.json"
    backup_path.write_bytes(COMPANIES_FILE.read_bytes())
    return backup_path


def count_companies() -> int:
    if not COMPANIES_FILE.exists():
        return 0
    return len(json.loads(COMPANIES_FILE.read_text(encoding="utf-8")))


def run_lead_extraction(fair_url: str, fair_name: str, append: bool = False):
    """
    Verilen URL'den firma listesini scrape et, sonra Bing + Google Maps
    enricher'ı çalıştır. UI'da canlı log gösterir.
    Her çalıştırmadan ÖNCE otomatik backup alır.
    Append modunda enricher --only-empty kullanır (mevcut firmaları atlar).
    """
    log_box = st.empty()
    progress = st.empty()

    # 0. OTOMATİK YEDEK
    n_before = count_companies()
    backup_path = backup_companies()
    if backup_path:
        st.toast(f"💾 {n_before} lead yedeklendi: {backup_path.name}", icon="💾")

    # 1. Scraper
    log_box.info(f"⏳ **{fair_name}** için katılımcı listesi çekiliyor...\n\n`{fair_url}`")
    cmd = [sys.executable, "scraper.py", "--url", fair_url, "--name", fair_name]
    if append:
        cmd.append("--append")
    proc1 = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    if proc1.returncode != 0:
        log_box.error(f"Scraper hata:\n```\n{proc1.stderr[-500:]}\n```")
        return

    n_after = count_companies()
    n_new = n_after - n_before if append else n_after

    # 0 yeni firma → enricher çalıştırma, kullanıcıyı uyar
    if n_new == 0:
        log_box.warning(
            f"⚠️ Bu URL'de **yeni firma bulunamadı**. "
            f"({n_after} firma listede ama hepsi zaten kayıtlı.)\n\n"
            f"**Olası sebepler:**\n"
            f"- URL fuarın katılımcı listesi değil (örn. düzenleyici firmanın ana sayfası)\n"
            f"- Sayfada katılımcı listesi var ama farklı yapıda (manuel URL girin)\n"
            f"- Fuar henüz katılımcı listesini yayınlamadı\n\n"
            f"💡 **Çözüm:** Fuarın kendi sitesine git, 'Katılımcılar' / 'Exhibitors' "
            f"sayfasının URL'ini kopyalayıp yukarıdaki kutuya yapıştır."
        )
        return

    log_box.success(
        f"✓ Scraper bitti: **{n_new} yeni firma** eklendi "
        f"(toplam: {n_after}). Enricher başlıyor..."
    )

    # 2. Enricher — append modunda sadece yeni/eksik firmaları işle
    enrich_cmd = [sys.executable, "enricher.py"]
    if append:
        enrich_cmd.append("--only-empty")
        target_n = n_new
    else:
        target_n = n_after

    progress.info(
        f"⏳ Enricher çalışıyor — **{target_n} firma** için Bing + Google Maps... "
        f"(yaklaşık ~{target_n * 12 / 60:.0f} dakika)"
    )
    proc2 = subprocess.run(enrich_cmd, capture_output=True, text=True, timeout=7200)
    if proc2.returncode != 0:
        progress.error(f"Enricher hata:\n```\n{proc2.stderr[-500:]}\n```")
        return

    progress.success(
        f"✅ **{fair_name}** için lead bulma tamamlandı! "
        f"📊 sekmesinden {n_new} yeni firmayı görebilirsiniz."
    )
    st.cache_data.clear()


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Fuarlar")
    return buf.getvalue()


# ── Üst başlık ───────────────────────────────────────────────────────────────

col_title, col_btn = st.columns([5, 1])
with col_title:
    st.title("🏭 Reyart Fuar Lead Sistemi")
    st.caption("Türkiye fuar takvimi → katılımcı firmalar → iletişim bilgileri")
with col_btn:
    st.write("")  # boşluk
    if st.button("🔄 Takvimi Güncelle", use_container_width=True):
        refresh_calendar()
        st.rerun()


# ── Sekme: Fuarlar | Lead'ler ────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs(["📅 Fuar Takvimi", "🎯 Lead Bul", "📊 Mevcut Lead'ler"])

# =============================================================================
# TAB 1 — Fuar Takvimi
# =============================================================================
with tab1:
    fairs = load_fairs()

    if fairs.empty:
        st.warning("Henüz fuar takvimi yok. **'Takvimi Güncelle'** butonuna basın.")
    else:
        # ── Filtreler ──
        c1, c2, c3, c4 = st.columns([2, 2, 2, 3])

        with c1:
            sectors = ["(Tümü)"] + sorted(fairs["sector"].dropna().unique().tolist())
            sel_sector = st.selectbox("Sektör", sectors)

        with c2:
            cities = ["(Tümü)"] + sorted(
                [c for c in fairs["city"].dropna().unique() if c]
            )
            sel_city = st.selectbox("Şehir", cities)

        with c3:
            today = date.today()
            min_d = fairs["start_dt"].min().date() if not fairs["start_dt"].isna().all() else today
            max_d = fairs["end_dt"].max().date() if not fairs["end_dt"].isna().all() else today
            sel_date = st.date_input(
                "Tarih aralığı",
                value=(min_d, max_d),
                min_value=min_d,
                max_value=max_d,
                help="Varsayılan: tüm yıl. Geçmiş fuarların katılımcıları da değerli lead'lerdir.",
            )

        with c4:
            search = st.text_input("🔍 Fuar adı, konu, düzenleyici...", "")

        # ── Filtre uygula ──
        f = fairs.copy()
        if sel_sector != "(Tümü)":
            f = f[f["sector"] == sel_sector]
        if sel_city != "(Tümü)":
            f = f[f["city"] == sel_city]

        if isinstance(sel_date, tuple) and len(sel_date) == 2:
            d_start, d_end = sel_date
            f = f[
                (f["start_dt"].dt.date >= d_start) &
                (f["start_dt"].dt.date <= d_end)
            ]

        if search:
            search_lower = search.lower()
            mask = (
                f["name"].str.lower().str.contains(search_lower, na=False)
                | f.get("topic", pd.Series("", index=f.index)).fillna("").str.lower().str.contains(search_lower, na=False)
                | f.get("organizer", pd.Series("", index=f.index)).fillna("").str.lower().str.contains(search_lower, na=False)
            )
            f = f[mask]

        # ── Üst metrikler ──
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Toplam Fuar", len(fairs))
        m2.metric("Filtrelenmiş", len(f))
        m3.metric("Şehir Sayısı", f["city"].nunique() if not f.empty else 0)
        with_participants = f[f.get("participants_url", "").astype(str).str.len() > 0]
        m4.metric("Katılımcı Listesi Var", len(with_participants))

        st.divider()

        # ── Tablo ──
        if f.empty:
            st.info("Filtrelere uyan fuar yok.")
        else:
            display = f[[
                "name", "start_date", "end_date", "city", "sector",
                "organizer", "url", "participants_url",
            ]].copy()
            display.columns = [
                "Fuar Adı", "Başlangıç", "Bitiş", "Şehir", "Sektör",
                "Düzenleyici", "Web", "Katılımcı Listesi",
            ]
            display = display.sort_values("Başlangıç")

            st.dataframe(
                display,
                use_container_width=True,
                height=520,
                hide_index=True,
                column_config={
                    "Web": st.column_config.LinkColumn(),
                    "Katılımcı Listesi": st.column_config.LinkColumn(),
                },
            )

            # İndir butonu
            xls_bytes = to_excel_bytes(display)
            st.download_button(
                "📥 Excel İndir",
                data=xls_bytes,
                file_name=f"fuar_takvimi_{datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# =============================================================================
# TAB 2 — Lead Bul
# =============================================================================
with tab2:
    st.subheader("🎯 Belirli bir fuar için lead bul")
    st.caption(
        "Generic scraper otomatik çalışır: önce tanımlı katılımcı listesi denenir, "
        "yoksa fuarın resmi sitesinden firmalar çıkarılmaya çalışılır. "
        "Sonra Bing + Google Maps ile telefon/e-posta zenginleştirme yapılır."
    )

    fairs = load_fairs()
    if fairs.empty:
        st.warning("Önce fuar takvimini yükleyin.")
    else:
        # ── Fuar seçici ──
        st.write("### 1️⃣ Fuar Seç")
        col_a, col_b = st.columns([3, 1])
        with col_a:
            fair_options = []
            for _, r in fairs.sort_values("start_date").iterrows():
                pl_mark = "🟢 " if str(r.get("participants_url", "")).strip() else (
                    "🟡 " if str(r.get("url", "")).strip() else "🔴 "
                )
                label = f"{pl_mark}{r['name'][:60]} · {r.get('start_date','')} · {r.get('city','')}"
                fair_options.append((label, r))

            sel_label = st.selectbox(
                "Fuar (🟢 katılımcı listesi tanımlı, 🟡 sadece web sitesi var, 🔴 URL yok)",
                options=[o[0] for o in fair_options],
                index=0,
            )
            sel_row = next(o[1] for o in fair_options if o[0] == sel_label)

        with col_b:
            st.write("")
            append_mode = st.checkbox(
                "🔒 Mevcut lead'lere ekle (önerilen)",
                value=True,
                help="VARSAYILAN AÇIK: önceki firmalar korunur. "
                     "Kapatırsan eski liste silinir — dikkat! "
                     "Her durumda otomatik yedek alınır.",
            )

        # ── Seçilen fuar bilgileri ──
        st.write("### 2️⃣ Bilgileri Doğrula")
        info_cols = st.columns(4)
        info_cols[0].metric("Tarih", sel_row.get("start_date", "—"))
        info_cols[1].metric("Şehir", sel_row.get("city", "—"))
        info_cols[2].metric("Sektör", sel_row.get("sector", "—")[:20])
        info_cols[3].metric("Kaynak", sel_row.get("source", "—"))

        # URL seçimi
        pl_url = str(sel_row.get("participants_url", "") or "").strip()
        web_url = str(sel_row.get("url", "") or "").strip()

        url_to_use = st.text_input(
            "Katılımcı listesi / fuar resmi web sitesi URL'si",
            value=pl_url or web_url,
            help="Otomatik tahmin: katılımcı listesi varsa o, yoksa resmi web sitesi.",
        )

        # ── Çalıştır ──
        st.write("### 3️⃣ Çalıştır")
        st.caption(
            "Tahmini süre: ~10–60 dakika (firma sayısına bağlı). "
            "İşlem boyunca Streamlit'i kapatmayın."
        )
        if st.button(
            "🚀 LEAD BUL",
            type="primary",
            disabled=not url_to_use,
            use_container_width=True,
        ):
            run_lead_extraction(url_to_use, sel_row["name"], append=append_mode)

# =============================================================================
# TAB 3 — CRM (Günlük Arama Paneli)
# =============================================================================

STATUS_OPTIONS = [
    "⬜ Aranmadı",
    "📞 Aradım - cevap yok",
    "📞 Aradım - tekrar arayacak",
    "✅ İlgileniyor",
    "💰 Teklif gönderildi",
    "🤝 Anlaşma",
    "❌ İlgilenmiyor",
]


def save_lead_update(lead_name: str, updates: dict):
    """Tek bir firma kaydını companies.json'da güncelle."""
    if not COMPANIES_FILE.exists():
        return False
    data = json.loads(COMPANIES_FILE.read_text(encoding="utf-8"))
    for c in data:
        if c.get("name") == lead_name:
            c.update(updates)
            # Çağrı kaydı tut
            if "status" in updates and updates["status"] != "⬜ Aranmadı":
                c.setdefault("call_log", [])
                c["call_log"].append({
                    "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "status": updates.get("status", ""),
                    "note": updates.get("note", ""),
                })
                c["last_call_date"] = datetime.now().strftime("%Y-%m-%d")
            break
    COMPANIES_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return True


def render_lead_card(lead: dict, idx: int):
    """Tek bir firma için kart + genişletilebilir düzenleme paneli."""
    name = lead.get("name", "")
    phone = lead.get("phone", "")
    priority = lead.get("priority", "⭐")
    status = lead.get("status", "⬜ Aranmadı")
    fair = lead.get("fair", "")

    # Renk kodu
    pri_emoji = {"⭐⭐⭐": "🔥", "⭐⭐": "💎", "⭐": "📌"}.get(priority, "📌")

    # Başlık satırı
    header = f"{pri_emoji} **{name[:60]}**  ·  📞 `{phone}`  ·  {status}"

    with st.expander(header, expanded=False):
        # Bilgi satırı
        info_cols = st.columns([2, 2, 2, 2])
        info_cols[0].markdown(f"**Öncelik:** {priority}")
        info_cols[1].markdown(f"**Fuar:** {fair[:25]}")
        info_cols[2].markdown(f"**Şehir:** {lead.get('country','—')}")
        info_cols[3].markdown(f"**Son arama:** {lead.get('last_call_date','—')}")

        # İletişim hızlı linkler
        link_cols = st.columns(4)
        if phone:
            phone_clean = re.sub(r"[^\d+]", "", phone)
            link_cols[0].markdown(f"📞 [Ara: {phone}](tel:{phone_clean})")
            link_cols[1].markdown(
                f"💬 [WhatsApp](https://wa.me/{phone_clean.lstrip('+')})"
            )
        if lead.get("email"):
            link_cols[2].markdown(f"📧 [Mail At](mailto:{lead['email']})")
        if lead.get("website"):
            link_cols[3].markdown(f"🌐 [Web]({lead['website']})")

        st.markdown("---")

        # Durum güncelleme formu
        with st.form(key=f"form_{idx}_{name[:20]}"):
            new_status = st.selectbox(
                "Durum güncelle",
                options=STATUS_OPTIONS,
                index=STATUS_OPTIONS.index(status) if status in STATUS_OPTIONS else 0,
                key=f"status_{idx}",
            )
            new_note = st.text_area(
                "Not (görüşme özeti, sonraki adım, vs.)",
                value=lead.get("note", ""),
                key=f"note_{idx}",
                height=80,
            )
            save_col, _ = st.columns([1, 4])
            if save_col.form_submit_button("💾 Kaydet", type="primary"):
                save_lead_update(name, {
                    "status": new_status,
                    "note": new_note,
                })
                st.cache_data.clear()
                st.success(f"✅ {name} güncellendi!")
                st.rerun()

        # Çağrı geçmişi
        log = lead.get("call_log", [])
        if log:
            st.caption("**Geçmiş:**")
            for entry in reversed(log[-5:]):
                st.caption(
                    f"📅 {entry.get('date','')}  ·  {entry.get('status','')}  "
                    f"·  _{entry.get('note','')[:80]}_"
                )


with tab3:
    st.subheader("🎯 CRM — Günlük Arama Paneli")
    leads = load_leads()

    if leads.empty:
        st.info("Henüz lead yok.")
    else:
        # ── Üst metrikler ──
        m1, m2, m3, m4, m5 = st.columns(5)
        sen_q = leads[leads.get("assigned_to", "") == "Sen"]
        ortak_q = leads[leads.get("assigned_to", "") == "Ortak"]
        m1.metric("Toplam", len(leads))
        m2.metric("🔥 Sen", len(sen_q),
                  delta=f"{(sen_q['status']=='⬜ Aranmadı').sum()} aranmadı")
        m3.metric("💎 Ortak", len(ortak_q),
                  delta=f"{(ortak_q['status']=='⬜ Aranmadı').sum()} aranmadı")
        m4.metric("✅ İlgileniyor",
                  (leads["status"].astype(str).str.contains("İlgileniyor", na=False)).sum())
        m5.metric("🤝 Anlaşma",
                  (leads["status"].astype(str).str.contains("Anlaşma", na=False)).sum())

        st.divider()

        # ── Sub-tab'lar ──
        view_tab1, view_tab2, view_tab3 = st.tabs([
            "🔥 Bugün - Sen", "💎 Bugün - Ortak", "📊 Tüm Liste",
        ])

        # ── Sen ──
        with view_tab1:
            st.caption(f"Sana atanmış **{len(sen_q)} firma** (⭐⭐⭐ + telefon var)")

            # Filtre
            fc1, fc2 = st.columns([1, 3])
            with fc1:
                only_pending = st.checkbox(
                    "Sadece aranmamış", value=True, key="sen_pending"
                )
            with fc2:
                sen_search = st.text_input(
                    "🔍 Firma ara", "", key="sen_search",
                    placeholder="Bobcat, Komatsu, Hidromek...",
                )

            sen_list = sen_q.copy()
            if only_pending:
                sen_list = sen_list[sen_list["status"] == "⬜ Aranmadı"]
            if sen_search:
                sen_list = sen_list[
                    sen_list["name"].str.contains(sen_search, case=False, na=False)
                ]
            sen_list = sen_list.sort_values("priority", ascending=False)

            st.write(f"**Listede {len(sen_list)} firma** — sırayla aç ve durumu işaretle:")
            for idx, row in enumerate(sen_list.to_dict(orient="records")):
                render_lead_card(row, idx)

        # ── Ortak ──
        with view_tab2:
            st.caption(f"Ortağa atanmış **{len(ortak_q)} firma** (⭐⭐ + telefon var)")

            fc1, fc2 = st.columns([1, 3])
            with fc1:
                only_pending_o = st.checkbox(
                    "Sadece aranmamış", value=True, key="ortak_pending"
                )
            with fc2:
                ortak_search = st.text_input(
                    "🔍 Firma ara", "", key="ortak_search",
                )

            ortak_list = ortak_q.copy()
            if only_pending_o:
                ortak_list = ortak_list[ortak_list["status"] == "⬜ Aranmadı"]
            if ortak_search:
                ortak_list = ortak_list[
                    ortak_list["name"].str.contains(ortak_search, case=False, na=False)
                ]

            st.write(f"**Listede {len(ortak_list)} firma**")
            for idx, row in enumerate(ortak_list.to_dict(orient="records")):
                render_lead_card(row, 1000 + idx)

        # ── Tüm Liste ──
        with view_tab3:
            f1, f2, f3, f4 = st.columns(4)
            with f1:
                fairs_in = ["(Tümü)"] + sorted(leads["fair"].dropna().unique().tolist())
                sel_fair = st.selectbox("Fuar", fairs_in, key="all_fair")
            with f2:
                priorities = ["(Tümü)", "⭐⭐⭐", "⭐⭐", "⭐"]
                sel_pri = st.selectbox("Öncelik", priorities, key="all_pri")
            with f3:
                statuses = ["(Tümü)"] + STATUS_OPTIONS
                sel_status = st.selectbox("Durum", statuses, key="all_status")
            with f4:
                assignees = ["(Tümü)", "Sen", "Ortak", "(Atanmamış)"]
                sel_assignee = st.selectbox("Kime atandı", assignees, key="all_ass")

            search_all = st.text_input("🔍 Firma adı / e-posta / telefon", "", key="all_search")

            l = leads.copy()
            if sel_fair != "(Tümü)":
                l = l[l["fair"] == sel_fair]
            if sel_pri != "(Tümü)":
                l = l[l["priority"] == sel_pri]
            if sel_status != "(Tümü)":
                l = l[l["status"] == sel_status]
            if sel_assignee == "(Atanmamış)":
                l = l[l.get("assigned_to", "").astype(str).str.len() == 0]
            elif sel_assignee != "(Tümü)":
                l = l[l.get("assigned_to", "") == sel_assignee]
            if search_all:
                m = (
                    l["name"].astype(str).str.contains(search_all, case=False, na=False)
                    | l["phone"].astype(str).str.contains(search_all, case=False, na=False)
                    | l["email"].astype(str).str.contains(search_all, case=False, na=False)
                )
                l = l[m]

            st.write(f"**{len(l)} sonuç**")

            # Görüntüleme tablosu
            display_cols = ["assigned_to", "name", "fair", "phone", "email",
                            "priority", "status", "note"]
            display_cols = [c for c in display_cols if c in l.columns]
            display = l[display_cols].copy()
            display.columns = [
                {"assigned_to": "Atanan", "name": "Firma", "fair": "Fuar",
                 "phone": "Telefon", "email": "E-posta",
                 "priority": "Öncelik", "status": "Durum", "note": "Not"}.get(c, c)
                for c in display_cols
            ]
            st.dataframe(display, use_container_width=True, height=500, hide_index=True)

            # İndir
            if LEADS_FILE.exists():
                with open(LEADS_FILE, "rb") as fp:
                    st.download_button(
                        "📥 Tüm lead'leri Excel olarak indir",
                        data=fp.read(),
                        file_name=f"leads_{datetime.now():%Y%m%d_%H%M}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


# ── Alt bilgi ────────────────────────────────────────────────────────────────

st.divider()
st.caption(
    "🛠️ **Reyart Lead Sistemi** — Kaynak: TOBB Resmi Fuar Takvimi · "
    "Enricher: Bing + Google Maps · Çıktı: Excel renk kodlu"
)
