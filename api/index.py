import os
import re
import csv
import json
import logging
from datetime import datetime, timedelta, timezone

from flask import Flask, request, jsonify
import gspread
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import AuthorizedSession, Request

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

# ── Config ────────────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID_1     = os.environ.get("SPREADSHEET_ID", "")
SPREADSHEET_ID_2     = os.environ.get("SPREADSHEET_ID_2", "")   # opsional
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON", "")
SHEET_NAME           = os.environ.get("SHEET_NAME", "Registrasi MUK")

# Google Drive pribadi (@gmail.com) via OAuth2 refresh token
GOOGLE_CLIENT_ID     = os.environ.get("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET", "")
GOOGLE_REFRESH_TOKEN = os.environ.get("GOOGLE_REFRESH_TOKEN", "")
DRIVE_FOLDER_ID      = os.environ.get("DRIVE_FOLDER_ID", "")
DRIVE_SCOPES         = ["https://www.googleapis.com/auth/drive"]

MAX_PDF_BYTES = 4 * 1024 * 1024   # batas body request Vercel ~4.5 MB
WITA = timezone(timedelta(hours=8))

HEADER = [
    "Timestamp",
    "Program Studi",
    "Dosen Diusulkan",
    "Judul Skema",
    "File Skema (PDF)",
]

FALLBACK_CSV = "/tmp/data_registrasi_muk.csv"


# ── Google Sheets ─────────────────────────────────────────────────────────────
def get_credentials():
    if SERVICE_ACCOUNT_JSON:
        info = json.loads(SERVICE_ACCOUNT_JSON)
        return Credentials.from_service_account_info(info, scopes=SCOPES)
    return Credentials.from_service_account_file("service_account.json", scopes=SCOPES)


def get_sheet(spreadsheet_id: str):
    gc = gspread.authorize(get_credentials())
    return gc.open_by_key(spreadsheet_id).worksheet(SHEET_NAME)


def ensure_header(sheet):
    if sheet.cell(1, 1).value != "Timestamp":
        sheet.insert_row(HEADER, 1)


def save_to_one_sheet(spreadsheet_id: str, row: list) -> bool:
    if not spreadsheet_id:
        return False
    try:
        sheet = get_sheet(spreadsheet_id)
        ensure_header(sheet)
        sheet.append_row(row, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        logging.error(f"[Google Sheets] Gagal ke {spreadsheet_id[:12]}...: {e}")
        return False


def save_to_sheets(row: list) -> bool:
    ok1 = save_to_one_sheet(SPREADSHEET_ID_1, row)
    ok2 = save_to_one_sheet(SPREADSHEET_ID_2, row)
    return ok1 or ok2


def save_to_csv(row: list):
    exists = os.path.isfile(FALLBACK_CSV)
    with open(FALLBACK_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(HEADER)
        w.writerow(row)


# ── Google Drive (akun pribadi, OAuth2) ───────────────────────────────────────
def drive_session():
    creds = UserCredentials(
        token=None,
        refresh_token=GOOGLE_REFRESH_TOKEN,
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=DRIVE_SCOPES,
    )
    creds.refresh(Request())
    return AuthorizedSession(creds)


def upload_to_drive(filename: str, data: bytes) -> str:
    """Upload PDF ke folder Drive pribadi, set 'siapa saja dengan link', return link."""
    sess = drive_session()
    meta = {"name": filename, "mimeType": "application/pdf"}
    if DRIVE_FOLDER_ID:
        meta["parents"] = [DRIVE_FOLDER_ID]

    boundary = "muk_boundary_7f3a9c"
    body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{json.dumps(meta)}\r\n--{boundary}\r\nContent-Type: application/pdf\r\n\r\n"
    ).encode("utf-8") + data + f"\r\n--{boundary}--".encode("utf-8")

    r = sess.post(
        "https://www.googleapis.com/upload/drive/v3/files"
        "?uploadType=multipart&fields=id,webViewLink&supportsAllDrives=true",
        data=body,
        headers={"Content-Type": f"multipart/related; boundary={boundary}"},
        timeout=60,
    )
    r.raise_for_status()
    info = r.json()

    sess.post(
        f"https://www.googleapis.com/drive/v3/files/{info['id']}/permissions?supportsAllDrives=true",
        json={"role": "reader", "type": "anyone"},
        timeout=30,
    ).raise_for_status()
    return info["webViewLink"]


# ── Format kolom ──────────────────────────────────────────────────────────────
def format_dosen(dosen_list: list) -> str:
    parts = []
    for d in dosen_list:
        nama = (d.get("nama") or "").strip()
        hp = (d.get("hp") or "").strip()
        if nama and hp:
            parts.append(f"{nama} ({hp})")
        elif nama:
            parts.append(nama)
    return " | ".join(parts)


def numbered(items: list) -> str:
    return "\n".join(f"{i}. {x}" for i, x in enumerate(items, 1))


# ── HTML ──────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Registrasi Workshop MUK</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet" />
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    :root {
      --navy:#0d2137; --gold:#c9972c; --gold-lt:#f0c96a; --cream:#faf7f2;
      --white:#ffffff; --gray:#6b7280; --gray-lt:#e5e7eb;
      --danger:#dc2626; --success:#16a34a; --radius:10px;
      --shadow:0 4px 24px rgba(13,33,55,.10);
    }
    body { font-family:'DM Sans',sans-serif; background:var(--cream); color:var(--navy); min-height:100vh; padding:0 0 60px; }
    .banner { background:var(--navy); color:var(--white); padding:36px 24px 32px; text-align:center; position:relative; overflow:hidden; }
    .banner::before { content:''; position:absolute; inset:0;
      background:repeating-linear-gradient(45deg,transparent,transparent 18px,rgba(201,151,44,.08) 18px,rgba(201,151,44,.08) 36px); }
    .banner-inner { position:relative; z-index:1; max-width:680px; margin:auto; }
    .banner h1 { font-family:'Playfair Display',serif; font-size:clamp(20px,4vw,28px); line-height:1.25; margin-bottom:10px; }
    .banner h1 span { color:var(--gold-lt); }
    .banner p { font-size:13px; opacity:.75; line-height:1.6; max-width:560px; margin:0 auto; }
    .card { background:var(--white); border-radius:var(--radius); box-shadow:var(--shadow); max-width:680px;
      margin:-20px auto 0; padding:36px 32px 40px; position:relative; z-index:2; }
    @media (max-width:600px){ .card { margin:-16px 12px 0; padding:24px 18px 32px; } }

    .section-label { font-size:11px; font-weight:600; letter-spacing:.08em; text-transform:uppercase; color:var(--gold);
      margin:28px 0 14px; display:flex; align-items:center; gap:8px; }
    .section-label:first-child { margin-top:0; }
    .section-label::after { content:''; flex:1; height:1px; background:var(--gray-lt); }
    .section-label small { font-weight:400; text-transform:none; letter-spacing:0; color:var(--gray); font-size:11px; }

    .field { margin-bottom:18px; }
    .field label { display:block; font-size:13px; font-weight:600; margin-bottom:6px; }
    .field label .req { color:var(--danger); margin-left:2px; }
    .field input[type="text"], .field input[type="tel"], .field select {
      width:100%; padding:10px 14px; border:1.5px solid var(--gray-lt); border-radius:var(--radius);
      font-family:inherit; font-size:14px; color:var(--navy); transition:border-color .2s, box-shadow .2s;
      outline:none; background:var(--white); }
    .field select { appearance:none; -webkit-appearance:none; padding-right:36px; cursor:pointer;
      background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%236b7280' stroke-width='2.5'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");
      background-repeat:no-repeat; background-position:right 12px center; }
    .field input:focus, .field select:focus { border-color:var(--gold); box-shadow:0 0 0 3px rgba(201,151,44,.14); }

    .dynamic-list { display:flex; flex-direction:column; gap:10px; }

    .dosen-row { display:grid; grid-template-columns:1fr 1fr auto; gap:8px; align-items:center;
      background:#f8fafc; border:1.5px solid var(--gray-lt); border-radius:var(--radius); padding:10px 12px; }
    .dosen-row input { padding:8px 10px; border:1.5px solid var(--gray-lt); border-radius:6px; font-family:inherit;
      font-size:13px; outline:none; background:var(--white); transition:border-color .2s; min-width:0; }
    .dosen-row input:focus { border-color:var(--gold); }
    .dosen-row input::placeholder { color:#9ca3af; font-size:12px; }
    @media (max-width:600px){ .dosen-row { grid-template-columns:1fr auto; } .dosen-row input:nth-child(2){ grid-column:1; } }

    .skema-row { display:grid; grid-template-columns:1fr auto; gap:8px; align-items:start;
      background:#f8fafc; border:1.5px solid var(--gray-lt); border-radius:var(--radius); padding:12px; }
    .skema-fields { display:flex; flex-direction:column; gap:8px; min-width:0; }
    .skema-fields input[type="text"] { padding:10px 14px; border:1.5px solid var(--gray-lt); border-radius:var(--radius);
      font-family:inherit; font-size:14px; outline:none; width:100%; background:var(--white); transition:border-color .2s; }
    .skema-fields input[type="text"]:focus { border-color:var(--gold); }
    .skema-fields input[type="file"] { font-family:inherit; font-size:12.5px; color:var(--gray); width:100%; }
    .skema-fields input[type="file"]::file-selector-button {
      border:1.5px solid var(--gold); background:rgba(201,151,44,.09); color:var(--gold); font-family:inherit;
      font-weight:600; font-size:12px; padding:6px 12px; border-radius:20px; cursor:pointer; margin-right:10px; }

    .btn-remove { width:30px; height:30px; border-radius:50%; border:1.5px solid #fecaca; background:#fff5f5; color:var(--danger);
      font-size:16px; cursor:pointer; display:flex; align-items:center; justify-content:center; flex-shrink:0; transition:background .15s; }
    .btn-remove:hover { background:#fee2e2; }
    .btn-remove:disabled { opacity:.3; cursor:not-allowed; }

    .btn-add { margin-top:8px; border:1.5px dashed var(--gold); background:transparent; color:var(--gold);
      padding:8px 16px; border-radius:var(--radius); font-family:inherit; font-size:13px; font-weight:600;
      cursor:pointer; transition:background .15s; display:inline-flex; align-items:center; gap:6px; }
    .btn-add:hover { background:rgba(201,151,44,.07); }

    .btn-ref { display:inline-flex; align-items:center; gap:6px; margin-top:6px; font-size:12px; font-weight:600; color:var(--gold);
      background:rgba(201,151,44,.09); border:1px solid rgba(201,151,44,.3); border-radius:20px; padding:5px 12px;
      text-decoration:none; transition:background .15s; }
    .btn-ref:hover { background:rgba(201,151,44,.18); }
    .btn-ref svg { width:13px; height:13px; }

    .btn-submit { margin-top:30px; width:100%; background:var(--navy); color:var(--white); border:none; border-radius:var(--radius);
      padding:14px; font-family:inherit; font-size:15px; font-weight:600; cursor:pointer;
      display:flex; align-items:center; justify-content:center; gap:8px; transition:background .2s, transform .1s; }
    .btn-submit:hover { background:#162f4f; }
    .btn-submit:active { transform:scale(.98); }
    .btn-submit:disabled { opacity:.6; cursor:not-allowed; }

    .alert { border-radius:var(--radius); padding:14px 16px; font-size:13px; line-height:1.5; margin-top:16px; display:none; }
    .alert.show { display:block; }
    .alert.danger { background:#fef2f2; border:1px solid #fecaca; color:var(--danger); }
    .alert ul { padding-left:18px; margin-top:4px; }

    .hint { font-size:11.5px; color:var(--gray); margin-top:4px; }

    .spinner { width:18px; height:18px; border:2.5px solid rgba(255,255,255,.4); border-top-color:#fff; border-radius:50%;
      animation:spin .7s linear infinite; display:none; }
    .btn-submit.loading .spinner { display:block; }
    .btn-submit.loading .btn-text { opacity:.7; }
    @keyframes spin { to { transform:rotate(360deg); } }

    .success-screen { display:none; text-align:center; padding:20px 0; }
    .success-screen.show { display:block; }
    .success-icon { width:64px; height:64px; border-radius:50%; background:#dcfce7; color:var(--success);
      display:flex; align-items:center; justify-content:center; margin:0 auto 20px; font-size:28px; }
    .success-screen h2 { font-family:'Playfair Display',serif; font-size:22px; margin-bottom:10px; }
    .success-screen p { font-size:14px; color:var(--gray); }
  </style>
</head>
<body>
<div class="banner">
  <div class="banner-inner">
    <h1>Workshop Penyusunan<br/><span>Materi Uji Kompetensi (MUK)</span></h1>
    <p>Isi formulir ini secara lengkap dan akurat. Unggah skema dalam format PDF.</p>
  </div>
</div>

<div class="card">
  <div class="success-screen" id="successScreen">
    <div class="success-icon">&#10003;</div>
    <h2>Pendaftaran Berhasil!</h2>
    <p>Data Anda telah tersimpan. Terima kasih telah mendaftar Workshop MUK.</p>
  </div>

  <div id="formWrapper">
    <div class="section-label">Data Prodi</div>

    <div class="field">
      <label>Program Studi <span class="req">*</span></label>
      <select id="prodi">
        <option value="">-- Pilih Program Studi --</option>
        <optgroup label="Fakultas Pertanian">
          <option>Budidaya Perairan</option>
          <option>Teknologi Pangan</option>
          <option>Peternakan</option>
          <option>Agribisnis</option>
          <option>Agroteknologi</option>
        </optgroup>
        <optgroup label="Fakultas Teknik">
          <option>Teknologi Informasi</option>
          <option>Teknik Geologi</option>
          <option>Perencanaan Wilayah dan Kota</option>
          <option>Teknik Arsitektur</option>
          <option>Teknik Lingkungan</option>
          <option>Teknik Sipil</option>
          <option>Teknik Kimia</option>
        </optgroup>
        <optgroup label="Fakultas Teknik Pertambangan">
          <option>Teknik Pertambangan</option>
        </optgroup>
        <optgroup label="Fakultas Ilmu Sosial &amp; Politik">
          <option>Sosiologi</option>
          <option>Administrasi Publik</option>
          <option>Ilmu Hubungan Internasional</option>
        </optgroup>
        <optgroup label="Fakultas Ekonomi &amp; Bisnis">
          <option>Ilmu Ekonomi Pembangunan</option>
          <option>Akuntansi</option>
          <option>Manajemen</option>
        </optgroup>
        <optgroup label="Fakultas Sastra &amp; Pendidikan">
          <option>Bahasa Mandarin</option>
          <option>Pendidikan Bahasa Indonesia</option>
        </optgroup>
        <optgroup label="Vokasi">
          <option>Teknik Mekatronika</option>
          <option>Teknik Listrik</option>
          <option>Perawatan dan Perbaikan Mesin</option>
          <option>Perhotelan</option>
          <option>Perpajakan</option>
        </optgroup>
      </select>
    </div>

    <div class="section-label">Dosen yang Diusulkan <small>(minimal 2 dosen)</small></div>
    <div class="dynamic-list" id="dosenList"></div>
    <button class="btn-add" type="button" onclick="tambahDosen()">+ Tambah Dosen</button>

    <div class="section-label">Skema Kompetensi <small>(judul + file PDF, maks. 4 MB per file)</small></div>
    <p class="hint" style="margin-bottom:10px;">Tulis judul skema, lalu unggah dokumen skemanya dalam format PDF. Referensi SKKNI resmi:</p>
    <a class="btn-ref" href="https://skkni.kemnaker.go.id/" target="_blank" rel="noopener">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M18 13v6a2 2 0 01-2 2H5a2 2 0 01-2-2V8a2 2 0 012-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
      Buka SKKNI Kemnaker (tab baru)
    </a>
    <div class="dynamic-list" id="skemaList" style="margin-top:14px;"></div>
    <button class="btn-add" type="button" onclick="tambahSkema()">+ Tambah Skema</button>

    <div class="alert danger" id="alertError"></div>

    <button class="btn-submit" id="btnSubmit" type="button" onclick="submitForm()">
      <span class="spinner"></span>
      <span class="btn-text">Kirim Pendaftaran</span>
    </button>
  </div>
</div>

<script>
  const MAX_PDF = 4 * 1024 * 1024;
  let dosenCount = 0, skemaCount = 0;

  function init() {
    tambahDosen(); tambahDosen();
    tambahSkema();
  }

  function tambahDosen() {
    dosenCount++;
    const id = dosenCount;
    const row = document.createElement('div');
    row.className = 'dosen-row'; row.id = `dosenRow${id}`;
    row.innerHTML = `
      <input type="text" placeholder="Nama Dosen ${id}" />
      <input type="tel" placeholder="No. HP / WA" />
      <button class="btn-remove" type="button" onclick="hapusRow('dosenRow${id}','dosenList','dosenRow',2)" title="Hapus">&#215;</button>`;
    document.getElementById('dosenList').appendChild(row);
    updateRemoveButtons('dosenList', 'dosenRow', 2);
  }

  function tambahSkema() {
    skemaCount++;
    const id = skemaCount;
    const row = document.createElement('div');
    row.className = 'skema-row'; row.id = `skemaRow${id}`;
    row.innerHTML = `
      <div class="skema-fields">
        <input type="text" placeholder="Judul skema, cth. Okupasi Pemrogram Junior" />
        <input type="file" accept="application/pdf,.pdf" />
      </div>
      <button class="btn-remove" type="button" onclick="hapusRow('skemaRow${id}','skemaList','skemaRow',1)" title="Hapus">&#215;</button>`;
    document.getElementById('skemaList').appendChild(row);
    updateRemoveButtons('skemaList', 'skemaRow', 1);
  }

  function hapusRow(rowId, listId, prefix, min) {
    const row = document.getElementById(rowId);
    if (row) row.remove();
    updateRemoveButtons(listId, prefix, min);
  }

  function updateRemoveButtons(listId, prefix, min) {
    const rows = document.getElementById(listId).querySelectorAll(`[id^="${prefix}"]`);
    rows.forEach(r => { const b = r.querySelector('.btn-remove'); if (b) b.disabled = rows.length <= min; });
  }

  function collectDosen() {
    return [...document.querySelectorAll('#dosenList .dosen-row')].map(row => {
      const i = row.querySelectorAll('input');
      return { nama: (i[0].value || '').trim(), hp: (i[1].value || '').trim() };
    }).filter(d => d.nama);
  }

  function collectSkema() {
    return [...document.querySelectorAll('#skemaList .skema-row')].map(row => ({
      judul: (row.querySelector('input[type="text"]').value || '').trim(),
      file: row.querySelector('input[type="file"]').files[0] || null,
    }));
  }

  function showErrors(list) {
    const el = document.getElementById('alertError');
    el.className = 'alert danger show';
    el.innerHTML = '<strong>Mohon perbaiki:</strong><ul>' + list.map(e => `<li>${e}</li>`).join('') + '</ul>';
    el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  async function uploadPdf(file) {
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetch('/upload', { method: 'POST', body: fd });
    const json = await res.json();
    if (!json.success) throw new Error((json.errors || ['Upload gagal.']).join(' '));
    return json.url;
  }

  async function submitForm() {
    const btn = document.getElementById('btnSubmit');
    const alertEl = document.getElementById('alertError');
    alertEl.className = 'alert'; alertEl.innerHTML = '';

    const payload = {
      prodi: document.getElementById('prodi').value,
      dosen: collectDosen(),
    };
    const skema = collectSkema();

    const errs = [];
    if (!payload.prodi) errs.push('Program Studi wajib dipilih.');
    if (payload.dosen.length < 2) errs.push('Minimal 2 dosen harus diisi.');
    if (payload.dosen.some(d => !d.hp)) errs.push('No. HP setiap dosen wajib diisi.');
    skema.forEach((s, i) => {
      const n = i + 1;
      if (!s.judul) errs.push(`Judul skema ${n} wajib diisi.`);
      if (!s.file) errs.push(`File PDF skema ${n} wajib diunggah.`);
      else if (!/\.pdf$/i.test(s.file.name)) errs.push(`File skema ${n} harus berformat PDF.`);
      else if (s.file.size > MAX_PDF) errs.push(`File skema ${n} melebihi 4 MB.`);
    });
    if (errs.length) return showErrors(errs);

    btn.disabled = true; btn.classList.add('loading');
    const txt = btn.querySelector('.btn-text');

    try {
      payload.skema = [];
      for (let i = 0; i < skema.length; i++) {
        txt.textContent = `Mengunggah skema ${i + 1} dari ${skema.length}...`;
        const url = await uploadPdf(skema[i].file);
        payload.skema.push({ judul: skema[i].judul, url });
      }
      txt.textContent = 'Menyimpan...';
      const res = await fetch('/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const json = await res.json();
      if (json.success) {
        document.getElementById('formWrapper').style.display = 'none';
        document.getElementById('successScreen').classList.add('show');
        window.scrollTo({ top: 0, behavior: 'smooth' });
      } else {
        showErrors(json.errors || ['Terjadi kesalahan.']);
      }
    } catch (e) {
      showErrors([e.message || 'Gagal terhubung ke server. Periksa koneksi Anda.']);
    } finally {
      btn.disabled = false; btn.classList.remove('loading');
      txt.textContent = 'Kirim Pendaftaran';
    }
  }

  init();
</script>
</body>
</html>"""


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET"])
def index():
    return HTML, 200, {"Content-Type": "text/html; charset=utf-8"}


@app.route("/upload", methods=["POST"])
def upload():
    """Upload satu file PDF ke Google Drive pribadi, kembalikan link-nya."""
    try:
        f = request.files.get("file")
        if not f or not f.filename:
            return jsonify({"success": False, "errors": ["File tidak ditemukan."]}), 422
        data = f.read()
        if not data.startswith(b"%PDF"):
            return jsonify({"success": False, "errors": ["File harus berformat PDF."]}), 422
        if len(data) > MAX_PDF_BYTES:
            return jsonify({"success": False, "errors": ["Ukuran PDF maksimal 4 MB."]}), 422
        if not (GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET and GOOGLE_REFRESH_TOKEN):
            return jsonify({"success": False, "errors": ["Kredensial Google Drive belum diatur."]}), 500

        stamp = datetime.now(WITA).strftime("%Y%m%d-%H%M%S")
        base = re.sub(r"[^A-Za-z0-9._-]", "_", f.filename)[:80]
        url = upload_to_drive(f"{stamp}-{base}", data)
        return jsonify({"success": True, "url": url})
    except Exception as e:
        logging.exception("Error di /upload")
        return jsonify({"success": False, "errors": [f"Upload gagal: {e}"]}), 500


@app.route("/submit", methods=["POST"])
def submit():
    try:
        data = request.get_json(force=True)

        prodi = (data.get("prodi") or "").strip()
        dosen_list = data.get("dosen", [])
        skema_list = data.get("skema", [])

        errors = []
        if not prodi:        errors.append("Program Studi wajib dipilih.")

        dosen_valid = [d for d in dosen_list if isinstance(d, dict) and (d.get("nama") or "").strip()]
        skema_valid = [
            s for s in skema_list
            if isinstance(s, dict) and (s.get("judul") or "").strip() and (s.get("url") or "").startswith("https://")
        ]
        if len(dosen_valid) < 2: errors.append("Minimal 2 dosen yang diusulkan.")
        if len(skema_valid) < 1: errors.append("Minimal 1 skema (judul + file PDF).")

        if errors:
            return jsonify({"success": False, "errors": errors}), 422

        row = [
            datetime.now(WITA).strftime("%Y-%m-%d %H:%M:%S"),
            prodi,
            format_dosen(dosen_valid),
            numbered([s["judul"].strip() for s in skema_valid]),
            numbered([s["url"].strip() for s in skema_valid]),
        ]

        saved = save_to_sheets(row)
        if not saved:
            save_to_csv(row)

        return jsonify({"success": True, "saved_to": "google_sheets" if saved else "csv_fallback"})

    except Exception as e:
        logging.exception("Error di /submit")
        return jsonify({"success": False, "errors": [str(e)]}), 500


if __name__ == "__main__":
    app.run(debug=True)
