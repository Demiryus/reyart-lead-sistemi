"""
Reyart Fuar Lead Sistemi — Masaüstü Uygulaması v2 (sıfırdan yeniden yazım)
Çalıştırma: python desktop_app_v2.py

pywebview + Alpine.js/Tailwind CDN arayüz. 3 sekme: Fuar Takvimi / Lead Bul / CRM.
scraper.py / enricher.py / fair_calendar.py subprocess ile çağrılır.
maps_only.py bu dosyanın kapsamı dışında, ayrı çalıştırılır.
"""

import json
import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path

import webview

# ── Yollar ───────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).resolve().parent

DATA_DIR = APP_DIR / "output"
FAIRS_JSON = DATA_DIR / "fairs.json"
COMPANIES_JSON = DATA_DIR / "companies.json"
LEADS_XLSX = DATA_DIR / "leads.xlsx"

CALL_STATUSES = [
    "⬜ Aranmadı",
    "📞 Aradım - cevap yok",
    "📞 Aradım - tekrar arayacak",
    "✅ İlgileniyor",
    "💰 Teklif gönderildi",
    "🤝 Anlaşma",
    "❌ İlgilenmiyor",
]


def python_interpreter() -> str:
    """Frozen exe içindeyken sistemde kurulu python.exe kullanılır (scraper.py
    ve enricher.py .exe'ye gömülmez); geliştirmede mevcut yorumlayıcı kullanılır."""
    if getattr(sys, "frozen", False):
        return "python"
    return sys.executable


def child_env() -> dict:
    """Windows'ta subprocess'in stdout/stderr encoding'i varsayılan sistem
    locale'idir (cp1252 gibi) ve Türkçe karakterleri (İ, ş, ı, ğ, ö, ü, ç)
    bozar / decode hatası verir. UTF-8'i zorlamak bunu çözer."""
    e = os.environ.copy()
    e["PYTHONIOENCODING"] = "utf-8"
    e["PYTHONUTF8"] = "1"
    return e


def run_subprocess(args, **kw):
    kw.setdefault("env", child_env())
    kw.setdefault("encoding", "utf-8")
    kw.setdefault("errors", "replace")
    return subprocess.run(args, **kw)


def load_json(path: Path, fallback):
    if not path.exists():
        return fallback
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return fallback


def save_json(path: Path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def leads_count() -> int:
    return len(load_json(COMPANIES_JSON, []))


def backup_leads():
    if not COMPANIES_JSON.exists():
        return None
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = DATA_DIR / f"companies_backup_{stamp}.json"
    dest.write_bytes(COMPANIES_JSON.read_bytes())
    return dest


class Api:
    def __init__(self):
        self.window = None

    def bind_window(self, window):
        self.window = window

    def _emit_log(self, text: str):
        if self.window:
            self.window.evaluate_js(f"window.onLeadLog({json.dumps(text)})")

    def _emit_finished(self, success: bool, text: str):
        if self.window:
            self.window.evaluate_js(
                f"window.onLeadFinished({json.dumps(success)}, {json.dumps(text)})"
            )

    # ── Fuar Takvimi ──────────────────────────────────────────────────────
    def get_fairs(self):
        return load_json(FAIRS_JSON, [])

    def refresh_calendar(self):
        try:
            proc = run_subprocess(
                [python_interpreter(), str(APP_DIR / "fair_calendar.py")],
                capture_output=True, text=True, timeout=300, cwd=str(APP_DIR),
            )
        except Exception as exc:
            return {"ok": False, "msg": f"Hata: {exc}"}
        if proc.returncode != 0:
            return {"ok": False, "msg": f"Hata: {proc.stderr[-500:]}"}
        return {"ok": True, "msg": "Fuar takvimi güncellendi!"}

    # ── Lead Bul ──────────────────────────────────────────────────────────
    def run_lead_extraction(self, fair_url: str, fair_name: str, append: bool = True):
        thread = threading.Thread(
            target=self._extract_leads_worker,
            args=(fair_url, fair_name, append),
            daemon=True,
        )
        thread.start()
        return {"started": True}

    def _extract_leads_worker(self, fair_url: str, fair_name: str, append: bool):
        before = leads_count()
        backup = backup_leads()
        if backup:
            self._emit_log(f"💾 {before} lead yedeklendi: {backup.name}")

        self._emit_log(f"⏳ {fair_name} için katılımcı listesi çekiliyor...\n{fair_url}")
        scraper_cmd = [
            python_interpreter(), str(APP_DIR / "scraper.py"),
            "--url", fair_url, "--name", fair_name,
        ]
        if append:
            scraper_cmd.append("--append")

        try:
            scraper_proc = run_subprocess(
                scraper_cmd, capture_output=True, text=True, timeout=600, cwd=str(APP_DIR),
            )
        except Exception as exc:
            self._emit_finished(False, f"Scraper başlatılamadı: {exc}")
            return
        if scraper_proc.returncode != 0:
            self._emit_finished(False, f"Scraper hata:\n{scraper_proc.stderr[-500:]}")
            return

        after = leads_count()
        added = after - before if append else after
        if added == 0:
            self._emit_finished(
                False,
                f"⚠️ Bu URL'de yeni firma bulunamadı. ({after} firma listede, hepsi zaten kayıtlı.) "
                "Fuarın 'Katılımcılar/Exhibitors' sayfasının URL'ini deneyin.",
            )
            return

        self._emit_log(f"✓ Scraper bitti: {added} yeni firma eklendi (toplam: {after}). Enricher başlıyor...")

        enricher_cmd = [python_interpreter(), str(APP_DIR / "enricher.py"), "--no-maps"]
        target = added
        if append:
            enricher_cmd.append("--only-empty")
        else:
            target = after

        self._emit_log(
            f"⏳ Enricher çalışıyor — {target} firma için Bing (~{target * 8 / 60:.0f} dk). "
            "Google Maps fallback için ayrı olarak maps_only.py çalıştırın."
        )
        try:
            enricher_proc = run_subprocess(
                enricher_cmd, capture_output=True, text=True, timeout=7200, cwd=str(APP_DIR),
            )
        except Exception as exc:
            self._emit_finished(False, f"Enricher başlatılamadı: {exc}")
            return
        if enricher_proc.returncode != 0:
            self._emit_finished(False, f"Enricher hata:\n{enricher_proc.stderr[-500:]}")
            return

        self._emit_finished(
            True,
            f"✅ {fair_name} için lead bulma tamamlandı! {added} yeni firma. CRM sekmesinden bakabilirsiniz.",
        )

    # ── CRM ───────────────────────────────────────────────────────────────
    def get_leads(self):
        return load_json(COMPANIES_JSON, [])

    def get_status_options(self):
        return CALL_STATUSES

    def save_lead_update(self, lead_name: str, status: str, note: str):
        if not COMPANIES_JSON.exists():
            return {"ok": False, "msg": "companies.json yok"}
        leads = load_json(COMPANIES_JSON, [])
        target = next((l for l in leads if l.get("name") == lead_name), None)
        if target is None:
            return {"ok": False, "msg": "Firma bulunamadı"}

        target["status"] = status
        target["note"] = note
        if status != "⬜ Aranmadı":
            target.setdefault("call_log", [])
            target["call_log"].append({
                "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "status": status,
                "note": note,
            })
            target["last_call_date"] = datetime.now().strftime("%Y-%m-%d")

        save_json(COMPANIES_JSON, leads)
        return {"ok": True}

    # ── Excel ─────────────────────────────────────────────────────────────
    def export_fairs_excel(self, rows):
        return self._save_rows_as_excel(rows, "fuar_takvimi")

    def export_leads_excel(self, rows):
        return self._save_rows_as_excel(rows, "leads")

    def _save_rows_as_excel(self, rows, name_prefix: str):
        try:
            import pandas as pd
        except ImportError:
            return {"ok": False, "msg": "pandas kurulu değil"}

        suggested = f"{name_prefix}_{datetime.now():%Y%m%d_%H%M}.xlsx"
        picked = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=suggested,
            file_types=("Excel Dosyası (*.xlsx)",),
        )
        if not picked:
            return {"ok": False, "msg": "İptal edildi"}
        dest = picked if isinstance(picked, str) else picked[0]
        pd.DataFrame(rows).to_excel(dest, index=False, sheet_name="Sayfa1")
        return {"ok": True, "msg": f"Kaydedildi: {dest}"}

    def download_existing_leads_xlsx(self):
        if not LEADS_XLSX.exists():
            return {"ok": False, "msg": "leads.xlsx bulunamadı"}
        suggested = f"leads_{datetime.now():%Y%m%d_%H%M}.xlsx"
        picked = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=suggested,
            file_types=("Excel Dosyası (*.xlsx)",),
        )
        if not picked:
            return {"ok": False, "msg": "İptal edildi"}
        dest = picked if isinstance(picked, str) else picked[0]
        Path(dest).write_bytes(LEADS_XLSX.read_bytes())
        return {"ok": True, "msg": f"Kaydedildi: {dest}"}


PAGE_HTML = r"""
<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<title>Reyart Fuar Lead Sistemi</title>
<script src="https://cdn.tailwindcss.com"></script>
<script defer src="https://cdn.jsdelivr.net/npm/alpinejs@3.x.x/dist/cdn.min.js"></script>
<style>
  :root{
    --bg:#050b14; --panel:#081523cc; --line:#0e2c40;
    --cyan:#00e5ff; --cyan-dim:#0a7f96; --teal:#19f0d8;
    --amber:#ffb300; --red:#ff4d5e; --txt:#9fd8e8; --txt-dim:#4d7d92;
    --card-bg:#081523; --card-border:#0e2c40;
  }
  body{font-family:'Rajdhani',sans-serif;background:radial-gradient(ellipse at 50% 38%, #0a1a2c 0%, var(--bg) 65%);color:var(--txt);margin:0}
  body::after{content:"";position:fixed;inset:0;pointer-events:none;z-index:9;background:repeating-linear-gradient(0deg,transparent 0 2px,#00e5ff05 2px 4px)}
  .card{transition:box-shadow .15s;background:var(--card-bg);border:1px solid var(--card-border);border-radius:12px}
  .card:hover{box-shadow:0 0 12px var(--cyan-dim)}
  ::-webkit-scrollbar{width:8px;height:8px}
  ::-webkit-scrollbar-thumb{background:var(--line);border-radius:4px}
</style>
</head>
<body class="bg-slate-100 text-slate-800">

<div x-data="leadApp()" x-init="init()" class="min-h-screen flex flex-col">

  <div class="bg-white border-b px-6 py-4 flex items-center justify-between shadow-sm">
    <div>
      <h1 class="text-2xl font-bold">🏭 Reyart Fuar Lead Sistemi</h1>
      <p class="text-sm text-slate-500">Türkiye fuar takvimi → katılımcı firmalar → iletişim bilgileri</p>
    </div>
    <button @click="refreshCalendar()" :disabled="calendarBusy"
      class="bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white px-4 py-2 rounded-lg font-medium">
      <span x-show="!calendarBusy">🔄 Takvimi Güncelle</span>
      <span x-show="calendarBusy">⏳ Güncelleniyor...</span>
    </button>
  </div>
  <div x-show="calendarMsg" x-text="calendarMsg" class="px-6 py-2 text-sm" :class="calendarOk ? 'bg-green-50 text-green-700' : 'bg-red-50 text-red-700'"></div>

  <div class="bg-white border-b px-6 flex gap-1">
    <button @click="activeTab='fairs'" :class="activeTab==='fairs' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500'"
      class="px-4 py-3 border-b-2 font-medium">📅 Fuar Takvimi</button>
    <button @click="activeTab='find'" :class="activeTab==='find' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500'"
      class="px-4 py-3 border-b-2 font-medium">🎯 Lead Bul</button>
    <button @click="activeTab='crm'" :class="activeTab==='crm' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500'"
      class="px-4 py-3 border-b-2 font-medium">📊 CRM / Mevcut Lead'ler</button>
  </div>

  <div class="flex-1 overflow-y-auto p-6">

    <!-- FUAR TAKVİMİ -->
    <div x-show="activeTab==='fairs'">
      <template x-if="fairs.length===0">
        <div class="bg-yellow-50 text-yellow-800 p-4 rounded-lg">Henüz fuar takvimi yok. "Takvimi Güncelle" butonuna basın.</div>
      </template>
      <template x-if="fairs.length>0">
        <div>
          <div class="grid grid-cols-4 gap-3 mb-4">
            <select x-model="sectorFilter" class="border rounded-lg px-3 py-2">
              <option value="">(Tümü) Sektör</option>
              <template x-for="s in sectorList" :key="s"><option :value="s" x-text="s"></option></template>
            </select>
            <select x-model="cityFilter" class="border rounded-lg px-3 py-2">
              <option value="">(Tümü) Şehir</option>
              <template x-for="c in cityList" :key="c"><option :value="c" x-text="c"></option></template>
            </select>
            <input x-model="fairSearch" placeholder="🔍 Fuar adı, konu, düzenleyici..." class="border rounded-lg px-3 py-2 col-span-2">
          </div>

          <div class="grid grid-cols-4 gap-3 mb-4">
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Toplam Fuar</div><div class="text-xl font-bold" x-text="fairs.length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Filtrelenmiş</div><div class="text-xl font-bold" x-text="visibleFairs.length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Şehir Sayısı</div><div class="text-xl font-bold" x-text="cityList.length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Katılımcı Listesi Var</div><div class="text-xl font-bold" x-text="visibleFairs.filter(f=>f.participants_url).length"></div></div>
          </div>

          <button @click="exportFairs()" class="mb-3 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm">📥 Excel İndir</button>

          <div class="bg-white rounded-lg shadow-sm overflow-hidden">
            <table class="w-full text-sm">
              <thead class="bg-slate-50 text-left">
                <tr>
                  <th class="p-2">Fuar Adı</th><th class="p-2">Başlangıç</th><th class="p-2">Bitiş</th>
                  <th class="p-2">Şehir</th><th class="p-2">Sektör</th><th class="p-2">Düzenleyici</th>
                  <th class="p-2">Web</th><th class="p-2">Katılımcı Listesi</th>
                </tr>
              </thead>
              <tbody>
                <template x-for="f in visibleFairs.slice(0,300)" :key="f.name+f.start_date">
                  <tr class="border-t hover:bg-slate-50">
                    <td class="p-2" x-text="f.name"></td>
                    <td class="p-2" x-text="f.start_date"></td>
                    <td class="p-2" x-text="f.end_date"></td>
                    <td class="p-2" x-text="f.city"></td>
                    <td class="p-2" x-text="f.sector"></td>
                    <td class="p-2" x-text="f.organizer"></td>
                    <td class="p-2"><a :href="f.url" class="text-blue-600 underline" target="_blank" x-show="f.url">Web</a></td>
                    <td class="p-2"><a :href="f.participants_url" class="text-blue-600 underline" target="_blank" x-show="f.participants_url">Katılımcılar</a></td>
                  </tr>
                </template>
              </tbody>
            </table>
          </div>
        </div>
      </template>
    </div>

    <!-- LEAD BUL -->
    <div x-show="activeTab==='find'">
      <h2 class="text-lg font-semibold mb-1">🎯 Belirli bir fuar için lead bul</h2>
      <p class="text-sm text-slate-500 mb-4">Önce tanımlı katılımcı listesi denenir, yoksa fuarın resmi sitesi kullanılır. Sonra Bing ile zenginleştirilir. Google Maps fallback ayrı araç: maps_only.py.</p>

      <template x-if="fairs.length===0">
        <div class="bg-yellow-50 text-yellow-800 p-4 rounded-lg">Önce fuar takvimini yükleyin.</div>
      </template>
      <template x-if="fairs.length>0">
        <div class="bg-white rounded-lg shadow-sm p-5 space-y-4 max-w-3xl">
          <div>
            <label class="block text-sm font-medium mb-1">1️⃣ Fuar Seç</label>
            <select x-model="selectedFairIndex" @change="onFairChosen()" class="border rounded-lg px-3 py-2 w-full">
              <template x-for="(f,i) in fairsByDate" :key="i">
                <option :value="i" x-text="(f.participants_url?'🟢 ':(f.url?'🟡 ':'🔴 ')) + f.name.slice(0,60) + ' · ' + f.start_date + ' · ' + f.city"></option>
              </template>
            </select>
            <p class="text-xs text-slate-400 mt-1">🟢 katılımcı listesi tanımlı · 🟡 sadece web sitesi · 🔴 URL yok</p>
          </div>

          <div class="grid grid-cols-4 gap-3 text-sm" x-show="chosenFair">
            <div><div class="text-slate-400">Tarih</div><div x-text="chosenFair?.start_date"></div></div>
            <div><div class="text-slate-400">Şehir</div><div x-text="chosenFair?.city"></div></div>
            <div><div class="text-slate-400">Sektör</div><div x-text="chosenFair?.sector"></div></div>
            <div><div class="text-slate-400">Kaynak</div><div x-text="chosenFair?.source"></div></div>
          </div>

          <div>
            <label class="block text-sm font-medium mb-1">2️⃣ Katılımcı listesi / fuar web sitesi URL'si</label>
            <input x-model="targetUrl" class="border rounded-lg px-3 py-2 w-full">
          </div>

          <label class="flex items-center gap-2 text-sm">
            <input type="checkbox" x-model="appendToExisting">
            🔒 Mevcut lead'lere ekle (önerilen — kapatırsan eski liste silinir, her durumda otomatik yedek alınır)
          </label>

          <button @click="startExtraction()" :disabled="!targetUrl || extractionBusy"
            class="bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white px-5 py-2.5 rounded-lg font-semibold w-full">
            <span x-show="!extractionBusy">🚀 LEAD BUL</span>
            <span x-show="extractionBusy">⏳ Çalışıyor... (kapatmayın)</span>
          </button>

          <div x-show="logLines.length" class="bg-slate-900 text-slate-100 rounded-lg p-3 text-xs font-mono max-h-72 overflow-y-auto whitespace-pre-wrap">
            <template x-for="(l,i) in logLines" :key="i"><div x-text="l"></div></template>
          </div>
        </div>
      </template>
    </div>

    <!-- CRM -->
    <div x-show="activeTab==='crm'">
      <h2 class="text-lg font-semibold mb-3">📊 CRM — Günlük Arama Paneli</h2>
      <template x-if="leads.length===0">
        <div class="bg-blue-50 text-blue-700 p-4 rounded-lg">Henüz lead yok.</div>
      </template>
      <template x-if="leads.length>0">
        <div>
          <div class="grid grid-cols-5 gap-3 mb-4">
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Toplam</div><div class="text-xl font-bold" x-text="leads.length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">⬜ Aranmadı</div><div class="text-xl font-bold" x-text="leads.filter(l=>l.status==='⬜ Aranmadı').length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">✅ İlgileniyor</div><div class="text-xl font-bold" x-text="leads.filter(l=>(l.status||'').includes('İlgileniyor')).length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">🤝 Anlaşma</div><div class="text-xl font-bold" x-text="leads.filter(l=>(l.status||'').includes('Anlaşma')).length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">❌ İlgilenmiyor</div><div class="text-xl font-bold" x-text="leads.filter(l=>(l.status||'').includes('İlgilenmiyor')).length"></div></div>
          </div>

          <div class="grid grid-cols-4 gap-3 mb-4">
            <select x-model="leadFairFilter" class="border rounded-lg px-3 py-2">
              <option value="">(Tümü) Fuar</option>
              <template x-for="fr in leadFairList" :key="fr"><option :value="fr" x-text="fr"></option></template>
            </select>
            <select x-model="leadStatusFilter" class="border rounded-lg px-3 py-2">
              <option value="">(Tümü) Durum</option>
              <template x-for="s in statusOptions" :key="s"><option :value="s" x-text="s"></option></template>
            </select>
            <label class="flex items-center gap-2 border rounded-lg px-3 py-2">
              <input type="checkbox" x-model="onlyPendingLeads"> Sadece aranmamış
            </label>
            <input x-model="leadSearch" placeholder="🔍 Firma / telefon / e-posta" class="border rounded-lg px-3 py-2">
          </div>

          <div class="flex gap-2 mb-3">
            <button @click="exportLeads()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm">📥 Filtrelenmiş Listeyi Excel İndir</button>
            <button @click="downloadLeadsXlsx()" class="bg-slate-600 hover:bg-slate-700 text-white px-4 py-2 rounded-lg text-sm">📥 leads.xlsx (renkli, enricher çıktısı) İndir</button>
            <span class="self-center text-sm text-slate-500" x-text="visibleLeads.length + ' sonuç'"></span>
          </div>

          <div class="space-y-2">
            <template x-for="(lead, idx) in visibleLeads" :key="lead.name+idx">
              <div class="bg-white rounded-lg shadow-sm card">
                <div class="p-3 flex items-center justify-between cursor-pointer" @click="toggleLeadCard(idx)">
                  <div class="flex items-center gap-3">
                    <span x-text="{'⭐⭐⭐':'🔥','⭐⭐':'💎','⭐':'📌'}[lead.priority] || '📌'"></span>
                    <span class="font-semibold" x-text="lead.name"></span>
                    <span class="text-slate-400 text-sm" x-text="lead.phone"></span>
                  </div>
                  <span class="text-sm px-2 py-1 rounded bg-slate-100" x-text="lead.status || '⬜ Aranmadı'"></span>
                </div>
                <div x-show="expandedCard===idx" class="border-t p-4 space-y-3 text-sm">
                  <div class="grid grid-cols-4 gap-3">
                    <div><b>Öncelik:</b> <span x-text="lead.priority"></span></div>
                    <div><b>Fuar:</b> <span x-text="lead.fair"></span></div>
                    <div><b>Şehir:</b> <span x-text="lead.country || '—'"></span></div>
                    <div><b>Son arama:</b> <span x-text="lead.last_call_date || '—'"></span></div>
                  </div>
                  <div class="flex gap-4">
                    <a x-show="lead.phone" :href="'tel:'+cleanPhone(lead.phone)" class="text-blue-600">📞 Ara</a>
                    <a x-show="lead.phone" :href="'https://wa.me/'+cleanPhone(lead.phone).replace('+','')" target="_blank" class="text-green-600">💬 WhatsApp</a>
                    <a x-show="lead.email" :href="'mailto:'+lead.email" class="text-blue-600">📧 Mail At</a>
                    <a x-show="lead.website" :href="lead.website" target="_blank" class="text-blue-600">🌐 Web</a>
                  </div>
                  <div class="grid grid-cols-2 gap-3">
                    <select x-model="lead._newStatus" class="border rounded-lg px-3 py-2">
                      <template x-for="s in statusOptions" :key="s"><option :value="s" x-text="s"></option></template>
                    </select>
                    <button @click="saveLeadUpdate(lead)" class="bg-blue-600 hover:bg-blue-700 text-white rounded-lg px-4 py-2">💾 Kaydet</button>
                  </div>
                  <textarea x-model="lead._newNote" placeholder="Not (görüşme özeti, sonraki adım...)" class="border rounded-lg px-3 py-2 w-full" rows="2"></textarea>
                  <template x-if="lead.call_log && lead.call_log.length">
                    <div class="text-xs text-slate-500 space-y-1">
                      <template x-for="entry in lead.call_log.slice(-5).reverse()">
                        <div x-text="'📅 '+entry.date+' · '+entry.status+' · '+(entry.note||'')"></div>
                      </template>
                    </div>
                  </template>
                </div>
              </div>
            </template>
          </div>
        </div>
      </template>
    </div>

  </div>

  <div class="bg-white border-t px-6 py-2 text-xs text-slate-400 text-center">
    🛠️ Reyart Lead Sistemi — Kaynak: TOBB Resmi Fuar Takvimi · Enricher: Bing (Maps ayrı: maps_only.py)
  </div>
</div>

<script>
function leadApp(){
  return {
    activeTab: 'fairs',
    fairs: [], leads: [], statusOptions: [],
    sectorFilter:'', cityFilter:'', fairSearch:'',
    selectedFairIndex: 0, targetUrl:'', appendToExisting:true,
    extractionBusy:false, logLines:[],
    calendarBusy:false, calendarMsg:'', calendarOk:true,
    leadFairFilter:'', leadStatusFilter:'', onlyPendingLeads:false, leadSearch:'',
    expandedCard:null,

    async init(){
      window.onLeadLog = (m)=>{ this.logLines.push(m); };
      window.onLeadFinished = (ok,msg)=>{
        this.logLines.push(msg);
        this.extractionBusy=false;
        if(ok){ this.reloadLeads(); }
      };
      this.fairs = await pywebview.api.get_fairs();
      this.leads = await pywebview.api.get_leads();
      this.statusOptions = await pywebview.api.get_status_options();
      this.leads.forEach(l=>{ l._newStatus = l.status || '⬜ Aranmadı'; l._newNote = l.note || ''; });
      if(this.fairs.length){ this.selectedFairIndex = 0; this.onFairChosen(); }
    },
    async reloadFairs(){ this.fairs = await pywebview.api.get_fairs(); },
    async reloadLeads(){
      this.leads = await pywebview.api.get_leads();
      this.leads.forEach(l=>{ l._newStatus = l.status || '⬜ Aranmadı'; l._newNote = l.note || ''; });
    },

    get sectorList(){ return [...new Set(this.fairs.map(f=>f.sector).filter(Boolean))].sort(); },
    get cityList(){ return [...new Set(this.fairs.map(f=>f.city).filter(Boolean))].sort(); },
    get visibleFairs(){
      return this.fairs.filter(f=>{
        if(this.sectorFilter && f.sector!==this.sectorFilter) return false;
        if(this.cityFilter && f.city!==this.cityFilter) return false;
        if(this.fairSearch){
          const q=this.fairSearch.toLowerCase();
          const hay=((f.name||'')+' '+(f.topic||'')+' '+(f.organizer||'')).toLowerCase();
          if(!hay.includes(q)) return false;
        }
        return true;
      });
    },
    get fairsByDate(){ return [...this.fairs].sort((a,b)=>(a.start_date||'').localeCompare(b.start_date||'')); },
    get chosenFair(){ return this.fairsByDate[this.selectedFairIndex]; },
    onFairChosen(){
      const f=this.chosenFair; if(!f) return;
      this.targetUrl = f.participants_url || f.url || '';
    },

    async refreshCalendar(){
      this.calendarBusy=true; this.calendarMsg='';
      const r = await pywebview.api.refresh_calendar();
      this.calendarBusy=false; this.calendarOk=r.ok; this.calendarMsg=r.msg;
      await this.reloadFairs();
    },

    async startExtraction(){
      this.extractionBusy=true; this.logLines=[];
      const f=this.chosenFair;
      await pywebview.api.run_lead_extraction(this.targetUrl, f.name, this.appendToExisting);
    },

    get leadFairList(){ return [...new Set(this.leads.map(l=>l.fair).filter(Boolean))].sort(); },
    get visibleLeads(){
      return this.leads.filter(l=>{
        if(this.leadFairFilter && l.fair!==this.leadFairFilter) return false;
        if(this.leadStatusFilter && l.status!==this.leadStatusFilter) return false;
        if(this.onlyPendingLeads && l.status && l.status!=='⬜ Aranmadı') return false;
        if(this.leadSearch){
          const q=this.leadSearch.toLowerCase();
          const hay=((l.name||'')+' '+(l.phone||'')+' '+(l.email||'')).toLowerCase();
          if(!hay.includes(q)) return false;
        }
        return true;
      });
    },
    toggleLeadCard(idx){ this.expandedCard = this.expandedCard===idx ? null : idx; },
    cleanPhone(p){ return (p||'').replace(/[^\d+]/g,''); },
    async saveLeadUpdate(lead){
      const r = await pywebview.api.save_lead_update(lead.name, lead._newStatus, lead._newNote);
      if(r.ok){ await this.reloadLeads(); alert('✅ '+lead.name+' güncellendi'); }
      else{ alert('Hata: '+(r.msg||'')); }
    },
    async exportFairs(){
      const r = await pywebview.api.export_fairs_excel(this.visibleFairs);
      alert(r.msg || (r.ok?'Kaydedildi':'İptal'));
    },
    async exportLeads(){
      const rows = this.visibleLeads.map(({_newStatus,_newNote,...rest})=>rest);
      const r = await pywebview.api.export_leads_excel(rows);
      alert(r.msg || (r.ok?'Kaydedildi':'İptal'));
    },
    async downloadLeadsXlsx(){
      const r = await pywebview.api.download_existing_leads_xlsx();
      alert(r.msg || (r.ok?'Kaydedildi':'İptal'));
    },
  }
}
</script>
</body>
</html>
"""


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    api = Api()
    window = webview.create_window(
        "Reyart Fuar Lead Sistemi",
        html=PAGE_HTML,
        js_api=api,
        width=1400,
        height=900,
        min_size=(1000, 700),
    )
    api.bind_window(window)
    webview.start()


if __name__ == "__main__":
    main()
