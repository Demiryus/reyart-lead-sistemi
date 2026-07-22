"""
Reyart Fuar Lead Sistemi — Masaüstü Uygulaması (Streamlit'siz)
Çalıştırma (geliştirme): python desktop_app.py
Paketleme (.exe): pyinstaller desktop_app.spec

pywebview + saf HTML/JS (Alpine.js + Tailwind CDN) arayüz.
Backend mantığı ui.py'den taşındı (scraper.py / enricher.py / fair_calendar.py
subprocess ile aynen çağrılıyor, companies.json / fairs.json şeması korundu).
"""

import json
import os
import re
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path

import webview

# ── Yollar ───────────────────────────────────────────────────────────────────
# PyInstaller ile paketlendiğinde script'in yanındaki değil, .exe'nin
# çalıştığı klasördeki output/ kullanılır (kullanıcı verisi exe ile taşınır).
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

OUTPUT_DIR = BASE_DIR / "output"
FAIRS_FILE = OUTPUT_DIR / "fairs.json"
COMPANIES_FILE = OUTPUT_DIR / "companies.json"
LEADS_FILE = OUTPUT_DIR / "leads.xlsx"

STATUS_OPTIONS = [
    "⬜ Aranmadı",
    "📞 Aradım - cevap yok",
    "📞 Aradım - tekrar arayacak",
    "✅ İlgileniyor",
    "💰 Teklif gönderildi",
    "🤝 Anlaşma",
    "❌ İlgilenmiyor",
]


def _py_exe():
    """Geliştirme modunda sys.executable python.exe'dir; frozen exe içinde
    scraper.py/enricher.py'yi ayrı bir python yorumlayıcısıyla çalıştırmak
    gerekir (sistemde kurulu python varsayılır)."""
    if getattr(sys, "frozen", False):
        return "python"
    return sys.executable


def _subprocess_env():
    """Windows'ta child process'in konsol/pipe encoding'i varsayılan olarak
    sistem locale'i (ör. cp1252) olur; bu Türkçe karakterleri (İ, ş, ı, ğ...)
    loglarken backslashreplace ile "\\u0130" gibi bozuk metne çevirir.
    UTF-8'e zorlayarak hem çökmeyi hem de bozuk log metnini önlüyoruz."""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env


def _run(cmd, **kwargs):
    kwargs.setdefault("env", _subprocess_env())
    kwargs.setdefault("encoding", "utf-8")
    kwargs.setdefault("errors", "replace")
    return subprocess.run(cmd, **kwargs)


# ── Yardımcılar ──────────────────────────────────────────────────────────────

def _read_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def count_companies() -> int:
    return len(_read_json(COMPANIES_FILE, []))


def backup_companies():
    if not COMPANIES_FILE.exists():
        return None
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = OUTPUT_DIR / f"companies_backup_{ts}.json"
    backup_path.write_bytes(COMPANIES_FILE.read_bytes())
    return backup_path


# ── JS köprüsü (Python API) ─────────────────────────────────────────────────

class Api:
    def __init__(self):
        self.window = None

    def set_window(self, window):
        self.window = window

    def _log(self, msg: str):
        if self.window:
            safe = json.dumps(msg)
            self.window.evaluate_js(f"window.appendLog({safe})")

    def _log_done(self, ok: bool, msg: str):
        if self.window:
            safe = json.dumps(msg)
            self.window.evaluate_js(f"window.leadRunDone({json.dumps(ok)}, {safe})")

    # ── Fuar Takvimi ──
    def get_fairs(self):
        return _read_json(FAIRS_FILE, [])

    def refresh_calendar(self):
        try:
            proc = _run(
                [_py_exe(), str(BASE_DIR / "fair_calendar.py")],
                capture_output=True, text=True, timeout=300, cwd=str(BASE_DIR),
            )
            if proc.returncode == 0:
                return {"ok": True, "msg": "Fuar takvimi güncellendi!"}
            return {"ok": False, "msg": f"Hata: {proc.stderr[-500:]}"}
        except Exception as e:
            return {"ok": False, "msg": f"Hata: {e}"}

    # ── Lead Bul ──
    def run_lead_extraction(self, fair_url: str, fair_name: str, append: bool = True):
        """Arka planda çalışır, JS tarafına appendLog/leadRunDone ile bildirir."""
        t = threading.Thread(
            target=self._run_lead_extraction_worker,
            args=(fair_url, fair_name, append),
            daemon=True,
        )
        t.start()
        return {"started": True}

    def _run_lead_extraction_worker(self, fair_url, fair_name, append):
        n_before = count_companies()
        backup_path = backup_companies()
        if backup_path:
            self._log(f"💾 {n_before} lead yedeklendi: {backup_path.name}")

        self._log(f"⏳ {fair_name} için katılımcı listesi çekiliyor...\n{fair_url}")
        cmd = [_py_exe(), str(BASE_DIR / "scraper.py"), "--url", fair_url, "--name", fair_name]
        if append:
            cmd.append("--append")
        try:
            proc1 = _run(cmd, capture_output=True, text=True, timeout=600, cwd=str(BASE_DIR))
        except Exception as e:
            self._log_done(False, f"Scraper başlatılamadı: {e}")
            return
        if proc1.returncode != 0:
            self._log_done(False, f"Scraper hata:\n{proc1.stderr[-500:]}")
            return

        n_after = count_companies()
        n_new = n_after - n_before if append else n_after

        if n_new == 0:
            self._log_done(
                False,
                f"⚠️ Bu URL'de yeni firma bulunamadı. ({n_after} firma listede, hepsi zaten kayıtlı.) "
                f"Fuarın 'Katılımcılar/Exhibitors' sayfasının URL'ini deneyin.",
            )
            return

        self._log(f"✓ Scraper bitti: {n_new} yeni firma eklendi (toplam: {n_after}). Enricher başlıyor...")

        enrich_cmd = [_py_exe(), str(BASE_DIR / "enricher.py"), "--no-maps"]
        if append:
            enrich_cmd.append("--only-empty")
            target_n = n_new
        else:
            target_n = n_after

        self._log(f"⏳ Enricher çalışıyor — {target_n} firma için Bing (~{target_n * 8 / 60:.0f} dk). Google Maps fallback için ayrı olarak maps_only.py çalıştırın.")
        try:
            proc2 = _run(enrich_cmd, capture_output=True, text=True, timeout=7200, cwd=str(BASE_DIR))
        except Exception as e:
            self._log_done(False, f"Enricher başlatılamadı: {e}")
            return
        if proc2.returncode != 0:
            self._log_done(False, f"Enricher hata:\n{proc2.stderr[-500:]}")
            return

        self._log_done(True, f"✅ {fair_name} için lead bulma tamamlandı! {n_new} yeni firma. CRM sekmesinden bakabilirsiniz.")

    # ── CRM ──
    def get_leads(self):
        return _read_json(COMPANIES_FILE, [])

    def get_status_options(self):
        return STATUS_OPTIONS

    def save_lead_update(self, lead_name: str, status: str, note: str):
        if not COMPANIES_FILE.exists():
            return {"ok": False, "msg": "companies.json yok"}
        data = _read_json(COMPANIES_FILE, [])
        found = False
        for c in data:
            if c.get("name") == lead_name:
                c["status"] = status
                c["note"] = note
                if status != "⬜ Aranmadı":
                    c.setdefault("call_log", [])
                    c["call_log"].append({
                        "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "status": status,
                        "note": note,
                    })
                    c["last_call_date"] = datetime.now().strftime("%Y-%m-%d")
                found = True
                break
        if not found:
            return {"ok": False, "msg": "Firma bulunamadı"}
        COMPANIES_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return {"ok": True}

    # ── Excel dışa aktar ──
    def export_fairs_excel(self, rows):
        return self._export_excel(rows, "fuar_takvimi")

    def export_leads_excel(self, rows):
        return self._export_excel(rows, "leads")

    def _export_excel(self, rows, prefix):
        try:
            import pandas as pd
        except ImportError:
            return {"ok": False, "msg": "pandas kurulu değil"}
        default_name = f"{prefix}_{datetime.now():%Y%m%d_%H%M}.xlsx"
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=default_name,
            file_types=("Excel Dosyası (*.xlsx)",),
        )
        if not result:
            return {"ok": False, "msg": "İptal edildi"}
        path = result if isinstance(result, str) else result[0]
        df = pd.DataFrame(rows)
        df.to_excel(path, index=False, sheet_name="Sayfa1")
        return {"ok": True, "msg": f"Kaydedildi: {path}"}

    def download_existing_leads_xlsx(self):
        if not LEADS_FILE.exists():
            return {"ok": False, "msg": "leads.xlsx bulunamadı"}
        default_name = f"leads_{datetime.now():%Y%m%d_%H%M}.xlsx"
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=default_name,
            file_types=("Excel Dosyası (*.xlsx)",),
        )
        if not result:
            return {"ok": False, "msg": "İptal edildi"}
        path = result if isinstance(result, str) else result[0]
        Path(path).write_bytes(LEADS_FILE.read_bytes())
        return {"ok": True, "msg": f"Kaydedildi: {path}"}


INDEX_HTML = r"""
<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<title>Reyart Fuar Lead Sistemi</title>
<script src="https://cdn.tailwindcss.com"></script>
<script defer src="https://cdn.jsdelivr.net/npm/alpinejs@3.x.x/dist/cdn.min.js"></script>
<style>
  body{font-family: 'Segoe UI', sans-serif;}
  .card{transition: box-shadow .15s;}
  .card:hover{box-shadow:0 2px 10px rgba(0,0,0,.12);}
  ::-webkit-scrollbar{width:8px;height:8px;}
  ::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:4px;}
</style>
</head>
<body class="bg-slate-100 text-slate-800">

<div x-data="app()" x-init="init()" class="min-h-screen flex flex-col">

  <!-- Üst başlık -->
  <div class="bg-white border-b px-6 py-4 flex items-center justify-between shadow-sm">
    <div>
      <h1 class="text-2xl font-bold">🏭 Reyart Fuar Lead Sistemi</h1>
      <p class="text-sm text-slate-500">Türkiye fuar takvimi → katılımcı firmalar → iletişim bilgileri</p>
    </div>
    <button @click="refreshCalendar()" :disabled="calendarLoading"
      class="bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white px-4 py-2 rounded-lg font-medium">
      <span x-show="!calendarLoading">🔄 Takvimi Güncelle</span>
      <span x-show="calendarLoading">⏳ Güncelleniyor...</span>
    </button>
  </div>
  <div x-show="calendarMsg" x-text="calendarMsg" class="px-6 py-2 text-sm" :class="calendarOk ? 'bg-green-50 text-green-700' : 'bg-red-50 text-red-700'"></div>

  <!-- Tabs -->
  <div class="bg-white border-b px-6 flex gap-1">
    <button @click="tab='fairs'" :class="tab==='fairs' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500'"
      class="px-4 py-3 border-b-2 font-medium">📅 Fuar Takvimi</button>
    <button @click="tab='find'" :class="tab==='find' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500'"
      class="px-4 py-3 border-b-2 font-medium">🎯 Lead Bul</button>
    <button @click="tab='crm'" :class="tab==='crm' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500'"
      class="px-4 py-3 border-b-2 font-medium">📊 CRM / Mevcut Lead'ler</button>
  </div>

  <div class="flex-1 overflow-y-auto p-6">

    <!-- ================= TAB 1: FAIRS ================= -->
    <div x-show="tab==='fairs'">
      <template x-if="fairs.length===0">
        <div class="bg-yellow-50 text-yellow-800 p-4 rounded-lg">Henüz fuar takvimi yok. "Takvimi Güncelle" butonuna basın.</div>
      </template>
      <template x-if="fairs.length>0">
        <div>
          <div class="grid grid-cols-4 gap-3 mb-4">
            <select x-model="fSector" class="border rounded-lg px-3 py-2">
              <option value="">(Tümü) Sektör</option>
              <template x-for="s in sectors" :key="s"><option :value="s" x-text="s"></option></template>
            </select>
            <select x-model="fCity" class="border rounded-lg px-3 py-2">
              <option value="">(Tümü) Şehir</option>
              <template x-for="c in cities" :key="c"><option :value="c" x-text="c"></option></template>
            </select>
            <input x-model="fSearch" placeholder="🔍 Fuar adı, konu, düzenleyici..." class="border rounded-lg px-3 py-2 col-span-2">
          </div>

          <div class="grid grid-cols-4 gap-3 mb-4">
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Toplam Fuar</div><div class="text-xl font-bold" x-text="fairs.length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Filtrelenmiş</div><div class="text-xl font-bold" x-text="filteredFairs.length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Şehir Sayısı</div><div class="text-xl font-bold" x-text="cities.length"></div></div>
            <div class="bg-white rounded-lg p-3 text-center shadow-sm"><div class="text-xs text-slate-500">Katılımcı Listesi Var</div><div class="text-xl font-bold" x-text="filteredFairs.filter(f=>f.participants_url).length"></div></div>
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
                <template x-for="f in filteredFairs.slice(0,300)" :key="f.name+f.start_date">
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

    <!-- ================= TAB 2: FIND ================= -->
    <div x-show="tab==='find'">
      <h2 class="text-lg font-semibold mb-1">🎯 Belirli bir fuar için lead bul</h2>
      <p class="text-sm text-slate-500 mb-4">Önce tanımlı katılımcı listesi denenir, yoksa fuarın resmi sitesi kullanılır. Sonra Bing ile zenginleştirilir. Google Maps fallback ayrı araç: maps_only.py.</p>

      <template x-if="fairs.length===0">
        <div class="bg-yellow-50 text-yellow-800 p-4 rounded-lg">Önce fuar takvimini yükleyin.</div>
      </template>
      <template x-if="fairs.length>0">
        <div class="bg-white rounded-lg shadow-sm p-5 space-y-4 max-w-3xl">
          <div>
            <label class="block text-sm font-medium mb-1">1️⃣ Fuar Seç</label>
            <select x-model="selFairIdx" @change="onFairSelect()" class="border rounded-lg px-3 py-2 w-full">
              <template x-for="(f,i) in sortedFairs" :key="i">
                <option :value="i" x-text="(f.participants_url?'🟢 ':(f.url?'🟡 ':'🔴 ')) + f.name.slice(0,60) + ' · ' + f.start_date + ' · ' + f.city"></option>
              </template>
            </select>
            <p class="text-xs text-slate-400 mt-1">🟢 katılımcı listesi tanımlı · 🟡 sadece web sitesi · 🔴 URL yok</p>
          </div>

          <div class="grid grid-cols-4 gap-3 text-sm" x-show="selFair">
            <div><div class="text-slate-400">Tarih</div><div x-text="selFair?.start_date"></div></div>
            <div><div class="text-slate-400">Şehir</div><div x-text="selFair?.city"></div></div>
            <div><div class="text-slate-400">Sektör</div><div x-text="selFair?.sector"></div></div>
            <div><div class="text-slate-400">Kaynak</div><div x-text="selFair?.source"></div></div>
          </div>

          <div>
            <label class="block text-sm font-medium mb-1">2️⃣ Katılımcı listesi / fuar web sitesi URL'si</label>
            <input x-model="urlToUse" class="border rounded-lg px-3 py-2 w-full">
          </div>

          <label class="flex items-center gap-2 text-sm">
            <input type="checkbox" x-model="appendMode">
            🔒 Mevcut lead'lere ekle (önerilen — kapatırsan eski liste silinir, her durumda otomatik yedek alınır)
          </label>

          <button @click="runLeadExtraction()" :disabled="!urlToUse || running"
            class="bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white px-5 py-2.5 rounded-lg font-semibold w-full">
            <span x-show="!running">🚀 LEAD BUL</span>
            <span x-show="running">⏳ Çalışıyor... (kapatmayın)</span>
          </button>

          <div x-show="logLines.length" class="bg-slate-900 text-slate-100 rounded-lg p-3 text-xs font-mono max-h-72 overflow-y-auto whitespace-pre-wrap">
            <template x-for="(l,i) in logLines" :key="i"><div x-text="l"></div></template>
          </div>
        </div>
      </template>
    </div>

    <!-- ================= TAB 3: CRM ================= -->
    <div x-show="tab==='crm'">
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
            <select x-model="lFair" class="border rounded-lg px-3 py-2">
              <option value="">(Tümü) Fuar</option>
              <template x-for="fr in leadFairs" :key="fr"><option :value="fr" x-text="fr"></option></template>
            </select>
            <select x-model="lStatus" class="border rounded-lg px-3 py-2">
              <option value="">(Tümü) Durum</option>
              <template x-for="s in statusOptions" :key="s"><option :value="s" x-text="s"></option></template>
            </select>
            <label class="flex items-center gap-2 border rounded-lg px-3 py-2">
              <input type="checkbox" x-model="lOnlyPending"> Sadece aranmamış
            </label>
            <input x-model="lSearch" placeholder="🔍 Firma / telefon / e-posta" class="border rounded-lg px-3 py-2">
          </div>

          <div class="flex gap-2 mb-3">
            <button @click="exportLeadsFromView()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm">📥 Filtrelenmiş Listeyi Excel İndir</button>
            <button @click="downloadExistingLeadsXlsx()" class="bg-slate-600 hover:bg-slate-700 text-white px-4 py-2 rounded-lg text-sm">📥 leads.xlsx (renkli, enricher çıktısı) İndir</button>
            <span class="self-center text-sm text-slate-500" x-text="filteredLeads.length + ' sonuç'"></span>
          </div>

          <div class="space-y-2">
            <template x-for="(lead, idx) in filteredLeads" :key="lead.name+idx">
              <div class="bg-white rounded-lg shadow-sm card">
                <div class="p-3 flex items-center justify-between cursor-pointer" @click="toggleCard(idx)">
                  <div class="flex items-center gap-3">
                    <span x-text="{'⭐⭐⭐':'🔥','⭐⭐':'💎','⭐':'📌'}[lead.priority] || '📌'"></span>
                    <span class="font-semibold" x-text="lead.name"></span>
                    <span class="text-slate-400 text-sm" x-text="lead.phone"></span>
                  </div>
                  <span class="text-sm px-2 py-1 rounded bg-slate-100" x-text="lead.status || '⬜ Aranmadı'"></span>
                </div>
                <div x-show="openCard===idx" class="border-t p-4 space-y-3 text-sm">
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
                    <button @click="saveLead(lead)" class="bg-blue-600 hover:bg-blue-700 text-white rounded-lg px-4 py-2">💾 Kaydet</button>
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
function app(){
  return {
    tab: 'fairs',
    fairs: [], leads: [], statusOptions: [],
    fSector:'', fCity:'', fSearch:'',
    selFairIdx: 0, urlToUse:'', appendMode:true,
    running:false, logLines:[],
    calendarLoading:false, calendarMsg:'', calendarOk:true,
    lFair:'', lStatus:'', lOnlyPending:false, lSearch:'',
    openCard:null,

    async init(){
      window.appendLog = (m)=>{ this.logLines.push(m); };
      window.leadRunDone = (ok,msg)=>{
        this.logLines.push(msg);
        this.running=false;
        if(ok){ this.reloadLeads(); }
      };
      this.fairs = await pywebview.api.get_fairs();
      this.leads = await pywebview.api.get_leads();
      this.statusOptions = await pywebview.api.get_status_options();
      this.leads.forEach(l=>{ l._newStatus = l.status || '⬜ Aranmadı'; l._newNote = l.note || ''; });
      if(this.fairs.length){ this.selFairIdx = 0; this.onFairSelect(); }
    },
    async reloadFairs(){ this.fairs = await pywebview.api.get_fairs(); },
    async reloadLeads(){
      this.leads = await pywebview.api.get_leads();
      this.leads.forEach(l=>{ l._newStatus = l.status || '⬜ Aranmadı'; l._newNote = l.note || ''; });
    },

    get sectors(){ return [...new Set(this.fairs.map(f=>f.sector).filter(Boolean))].sort(); },
    get cities(){ return [...new Set(this.fairs.map(f=>f.city).filter(Boolean))].sort(); },
    get filteredFairs(){
      return this.fairs.filter(f=>{
        if(this.fSector && f.sector!==this.fSector) return false;
        if(this.fCity && f.city!==this.fCity) return false;
        if(this.fSearch){
          const q=this.fSearch.toLowerCase();
          const hay=((f.name||'')+' '+(f.topic||'')+' '+(f.organizer||'')).toLowerCase();
          if(!hay.includes(q)) return false;
        }
        return true;
      });
    },
    get sortedFairs(){ return [...this.fairs].sort((a,b)=>(a.start_date||'').localeCompare(b.start_date||'')); },
    get selFair(){ return this.sortedFairs[this.selFairIdx]; },
    onFairSelect(){
      const f=this.selFair; if(!f) return;
      this.urlToUse = f.participants_url || f.url || '';
    },

    async refreshCalendar(){
      this.calendarLoading=true; this.calendarMsg='';
      const r = await pywebview.api.refresh_calendar();
      this.calendarLoading=false; this.calendarOk=r.ok; this.calendarMsg=r.msg;
      await this.reloadFairs();
    },

    async runLeadExtraction(){
      this.running=true; this.logLines=[];
      const f=this.selFair;
      await pywebview.api.run_lead_extraction(this.urlToUse, f.name, this.appendMode);
    },

    get leadFairs(){ return [...new Set(this.leads.map(l=>l.fair).filter(Boolean))].sort(); },
    get filteredLeads(){
      return this.leads.filter(l=>{
        if(this.lFair && l.fair!==this.lFair) return false;
        if(this.lStatus && l.status!==this.lStatus) return false;
        if(this.lOnlyPending && l.status && l.status!=='⬜ Aranmadı') return false;
        if(this.lSearch){
          const q=this.lSearch.toLowerCase();
          const hay=((l.name||'')+' '+(l.phone||'')+' '+(l.email||'')).toLowerCase();
          if(!hay.includes(q)) return false;
        }
        return true;
      });
    },
    toggleCard(idx){ this.openCard = this.openCard===idx ? null : idx; },
    cleanPhone(p){ return (p||'').replace(/[^\d+]/g,''); },
    async saveLead(lead){
      const r = await pywebview.api.save_lead_update(lead.name, lead._newStatus, lead._newNote);
      if(r.ok){ await this.reloadLeads(); alert('✅ '+lead.name+' güncellendi'); }
      else{ alert('Hata: '+(r.msg||'')); }
    },
    async exportFairs(){
      const r = await pywebview.api.export_fairs_excel(this.filteredFairs);
      alert(r.msg || (r.ok?'Kaydedildi':'İptal'));
    },
    async exportLeadsFromView(){
      const rows = this.filteredLeads.map(({_newStatus,_newNote,...rest})=>rest);
      const r = await pywebview.api.export_leads_excel(rows);
      alert(r.msg || (r.ok?'Kaydedildi':'İptal'));
    },
    async downloadExistingLeadsXlsx(){
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
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    api = Api()
    window = webview.create_window(
        "Reyart Fuar Lead Sistemi",
        html=INDEX_HTML,
        js_api=api,
        width=1400,
        height=900,
        min_size=(1000, 700),
    )
    api.set_window(window)
    webview.start()


if __name__ == "__main__":
    main()
