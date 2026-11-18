import os
import time
import io
import re
import csv
import json
import logging
from datetime import datetime, timedelta, timezone

from flask import Flask, request, jsonify
import gspread
import requests
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
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON") or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
SHEET_NAME           = os.environ.get("SHEET_NAME", "Registrasi MUK")

# Daftar dosen (dropdown) — kolom B = nama
DOSEN_SPREADSHEET_ID = os.environ.get("DOSEN_SPREADSHEET_ID", "1aiEKOyJQNCau-Yng9lDvlAoimWKRHbLcI6lExFnKmwQ")
DOSEN_SHEET_GID      = os.environ.get("DOSEN_SHEET_GID", "0")
DOSEN_CACHE_TTL      = 300   # detik

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
    if os.path.isfile("service_account.json"):
        return Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
    raise RuntimeError("SERVICE_ACCOUNT_JSON belum diatur di Environment Variables Vercel.")


def get_sheet(spreadsheet_id: str):
    gc = gspread.authorize(get_credentials())
    sh = gc.open_by_key(spreadsheet_id)
    try:
        return sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        logging.info(f"Tab '{SHEET_NAME}' tidak ada, membuat tab baru.")
        try:
            return sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(HEADER))
        except gspread.exceptions.APIError:
            # tab baru saja dibuat oleh pengiriman lain pada saat yang sama
            return sh.worksheet(SHEET_NAME)


def ensure_header(sheet):
    if sheet.cell(1, 1).value != "Timestamp":
        sheet.insert_row(HEADER, 1)


def save_to_one_sheet(spreadsheet_id: str, row: list):
    """Return (berhasil, pesan_error)."""
    if not spreadsheet_id:
        return None, ""          # tidak diisi -> dilewati
    try:
        sheet = get_sheet(spreadsheet_id)
        ensure_header(sheet)
        sheet.append_row(row, value_input_option="USER_ENTERED")
        return True, ""
    except Exception as e:
        logging.exception(f"[Google Sheets] Gagal ke {spreadsheet_id[:12]}...")
        detail = str(e)
        cause = getattr(e, "__cause__", None)
        resp = getattr(cause, "response", None)
        if resp is not None:
            detail += " | Google: " + resp.text[:350].replace("\n", " ")
        if isinstance(e, PermissionError):
            try:
                detail += f" | Share spreadsheet ke: {json.loads(SERVICE_ACCOUNT_JSON)['client_email']}"
            except Exception:
                pass
        return False, f"{spreadsheet_id[:8]}...: {type(e).__name__} {detail}"


def save_to_sheets(row: list):
    """Return (ada_yang_berhasil, daftar_error)."""
    results = [save_to_one_sheet(i, row) for i in (SPREADSHEET_ID_1, SPREADSHEET_ID_2)]
    ok_any = any(r[0] is True for r in results)
    errs = [r[1] for r in results if r[0] is False]
    return ok_any, errs


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
    if not r.ok:
        raise RuntimeError(f"Drive {r.status_code}: {r.text[:400]}")
    info = r.json()

    pr = sess.post(
        f"https://www.googleapis.com/drive/v3/files/{info['id']}/permissions?supportsAllDrives=true",
        json={"role": "reader", "type": "anyone"},
        timeout=30,
    )
    if not pr.ok:
        raise RuntimeError(f"Drive permissions {pr.status_code}: {pr.text[:400]}")
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


_DOSEN_CACHE = {"ts": 0.0, "names": []}


def _bersihkan_daftar(values: list) -> list:
    """values = isi kolom B (baris pertama = header). Return nama unik, urutan asli."""
    out, seen = [], set()
    for i, v in enumerate(values):
        nama = str(v or "").strip().rstrip(" ,")
        if not nama or (i == 0 and "nama" in nama.casefold()):
            continue
        k = _norm(nama)
        if k not in seen:
            seen.add(k)
            out.append(nama)
    return out


def load_daftar_dosen() -> list:
    """Ambil nama dosen dari kolom B spreadsheet daftar dosen (cache 5 menit)."""
    now = time.time()
    if _DOSEN_CACHE["names"] and now - _DOSEN_CACHE["ts"] < DOSEN_CACHE_TTL:
        return _DOSEN_CACHE["names"]

    names = []
    try:   # 1) lewat service account
        ws = gspread.authorize(get_credentials()).open_by_key(DOSEN_SPREADSHEET_ID).get_worksheet(0)
        names = _bersihkan_daftar(ws.col_values(2))
    except Exception:
        logging.warning("Daftar dosen via service account gagal, coba export CSV publik.", exc_info=True)
    if not names:   # 2) cadangan: export CSV (sheet dibagikan 'anyone with the link')
        try:
            r = requests.get(
                f"https://docs.google.com/spreadsheets/d/{DOSEN_SPREADSHEET_ID}/export",
                params={"format": "csv", "gid": DOSEN_SHEET_GID}, timeout=15,
            )
            r.raise_for_status()
            rows = list(csv.reader(io.StringIO(r.content.decode("utf-8-sig"))))
            names = _bersihkan_daftar([row[1] if len(row) > 1 else "" for row in rows])
        except Exception:
            logging.exception("Gagal mengambil daftar dosen")

    if names:
        _DOSEN_CACHE.update(ts=now, names=names)
        return names
    return _DOSEN_CACHE["names"]     # data lama bila ada, kalau tidak: []


def kanonikkan_dosen(dosen_list: list):
    """Cocokkan nama yang dikirim dengan daftar dosen. Return (list_baru, errors).
    Jika daftar tidak bisa dimuat, validasi dilewati."""
    daftar = load_daftar_dosen()
    if not daftar:
        return dosen_list, []
    peta = {_norm(n): n for n in daftar}
    baru, errors = [], []
    for d in dosen_list:
        nama = (d.get("nama") or "").strip() if isinstance(d, dict) else ""
        if not nama:
            continue
        kanon = peta.get(_norm(nama))
        if kanon is None:
            errors.append(f'Nama dosen "{nama}" tidak ada dalam daftar dosen. Pilih dari dropdown.')
        else:
            baru.append({**d, "nama": kanon})
    return baru, errors


def _norm(text: str) -> str:
    """Normalisasi untuk perbandingan: huruf kecil, spasi rapi."""
    return re.sub(r"\s+", " ", (text or "").strip()).casefold()


def _nama_dari_sel(cell: str) -> list:
    """'Budi, M.Si (+628123) | Ani (0812)' -> ['Budi, M.Si', 'Ani'] (tanpa nomor HP)."""
    out = []
    for part in (cell or "").split("|"):
        nama = re.sub(r"\s*\(\+?[0-9]*\)\s*$", "", part.strip())
        if nama:
            out.append(nama)
    return out


def cek_dosen_duplikat(prodi: str, dosen_list: list) -> list:
    """Return daftar pesan error bila ada nama dosen yang sudah diusulkan
    di Program Studi yang sama (baris lain di spreadsheet) atau dobel dalam form ini."""
    errors = []
    baru = [(d.get("nama") or "").strip() for d in dosen_list if isinstance(d, dict)]
    baru = [n for n in baru if n]

    # 1) dobel di dalam form yang sama
    seen = set()
    for n in baru:
        k = _norm(n)
        if k in seen:
            errors.append(f'Nama dosen "{n}" diisi lebih dari sekali pada form ini.')
        seen.add(k)

    # 2) sudah ada di baris lain pada prodi yang sama
    sid = SPREADSHEET_ID_1 or SPREADSHEET_ID_2
    if not sid:
        return errors
    try:
        sheet = get_sheet(sid)
        rows = sheet.get("B:C")          # kolom B = Program Studi, C = Dosen Diusulkan
    except Exception:
        logging.exception("Gagal membaca spreadsheet untuk cek dosen ganda")
        return errors

    ada = set()
    for r in rows:
        if len(r) < 2 or _norm(r[0]) != _norm(prodi):
            continue
        for n in _nama_dari_sel(r[1]):
            ada.add(_norm(n))

    done = set()
    for n in baru:
        k = _norm(n)
        if k in ada and k not in done:
            errors.append(f'Nama dosen "{n}" sudah ada (sudah diusulkan) di Program Studi {prodi}.')
            done.add(k)
    return errors


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
    /* ── Dropdown dosen (combobox) ── */
    .cb { position:relative; min-width:0; }
    .dosen-row .cb-input { width:100%; padding:9px 34px 9px 12px; cursor:pointer; text-overflow:ellipsis; }
    .cb.has-value .cb-input { font-weight:600; color:var(--navy); border-color:#d9c08a; background:#fffdf8; }
    .cb.open .cb-input { border-color:var(--gold); box-shadow:0 0 0 3px rgba(201,151,44,.15); cursor:text; }
    .cb-chev { position:absolute; right:11px; top:50%; width:16px; height:16px; margin-top:-8px; color:var(--gray);
      pointer-events:none; transition:transform .2s; }
    .cb.open .cb-chev { transform:rotate(180deg); color:var(--gold); }
    .cb-panel { position:absolute; left:0; top:calc(100% + 6px); z-index:60; width:max(100%, 380px); max-width:calc(100vw - 48px);
      background:var(--white); border:1px solid var(--gray-lt); border-radius:12px; overflow:hidden;
      box-shadow:0 12px 32px rgba(13,33,55,.18), 0 2px 6px rgba(13,33,55,.08);
      opacity:0; transform:translateY(-4px); pointer-events:none; visibility:hidden; transition:opacity .15s, transform .15s, visibility .15s; }
    .cb.open .cb-panel { opacity:1; transform:none; pointer-events:auto; visibility:visible; }
    .cb.up .cb-panel { top:auto; bottom:calc(100% + 6px); }
    @media (max-width:600px){ .cb-panel { width:100%; } }
    .cb-meta { padding:8px 14px; font-size:11px; letter-spacing:.04em; text-transform:uppercase; color:var(--gray);
      background:#f8fafc; border-bottom:1px solid var(--gray-lt); }
    .cb-list { max-height:248px; overflow-y:auto; padding:6px; overscroll-behavior:contain; }
    .cb-opt { display:flex; align-items:center; gap:10px; padding:8px 10px; border-radius:8px; cursor:pointer; }
    .cb-opt.is-active { background:#f4efe3; }
    .cb-opt.is-selected { background:#fbf4df; }
    .cb-opt.is-disabled { opacity:.45; cursor:not-allowed; }
    .cb-avatar { flex:none; width:32px; height:32px; border-radius:50%; background:var(--navy); color:var(--gold-lt);
      display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:600; letter-spacing:.02em; }
    .cb-opt.is-selected .cb-avatar { background:var(--gold); color:var(--white); }
    .cb-txt { min-width:0; flex:1; line-height:1.25; }
    .cb-name { font-size:13.5px; font-weight:600; color:var(--navy); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .cb-deg { font-size:11.5px; color:var(--gray); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .cb-opt mark { background:rgba(201,151,44,.28); color:inherit; border-radius:3px; padding:0 1px; }
    .cb-tag { flex:none; font-size:10.5px; font-weight:600; color:var(--gray); background:var(--gray-lt); border-radius:999px; padding:2px 8px; }
    .cb-check { flex:none; width:16px; height:16px; color:var(--gold); }
    .cb-empty { padding:22px 14px; text-align:center; font-size:13px; color:var(--gray); }


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
    #formWrapper input:disabled, #formWrapper select:disabled { opacity:.6; cursor:not-allowed; background:#eef0f3; }
    .btn-add:disabled { opacity:.5; cursor:not-allowed; }
    .btn-remove:disabled { cursor:not-allowed; }

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

  let DAFTAR_DOSEN = [];
  let DOSEN_STATE = 'loading';      // loading | ok | fail

  async function muatDosen() {
    try {
      const r = await fetch('/api/dosen');
      const j = await r.json();
      if (!j.success) throw new Error('gagal');
      DAFTAR_DOSEN = j.dosen;
      DOSEN_STATE = 'ok';
    } catch (e) {
      DOSEN_STATE = 'fail';
      showErrors(['Daftar dosen gagal dimuat. Muat ulang halaman.']);
    }
    document.querySelectorAll('.cb.open').forEach(cbRender);
  }

  /* ── Combobox dosen ── */
  const TITLE_RE = /^(dr|prof|ir|drs|h|hj)\.?$/i;
  const esc = t => String(t).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

  function pecahNama(n) {
    const i = n.indexOf(',');
    return i < 0 ? [n, ''] : [n.slice(0, i).trim(), n.slice(i + 1).trim()];
  }
  function inisial(n) {
    const w = pecahNama(n)[0].split(/\s+/).filter(x => x && !TITLE_RE.test(x));
    const a = (w[0] || n)[0] || '?', b = w.length > 1 ? w[1][0] : '';
    return (a + b).toUpperCase();
  }
  function hl(text, q) {
    const t = q.split(/\s+/).filter(Boolean).map(x => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    if (!t.length) return esc(text);
    return text.split(new RegExp('(' + t.join('|') + ')', 'i')).map((p, i) => i % 2 ? '<mark>' + esc(p) + '</mark>' : esc(p)).join('');
  }

  function cbInit(cb) {
    const input = cb.querySelector('.cb-input'), panel = cb.querySelector('.cb-panel');
    cb._st = { sel: '', active: -1, items: [] };

    input.addEventListener('focus', () => { cbOpen(cb); input.select(); });
    input.addEventListener('click', () => cbOpen(cb));
    input.addEventListener('input', () => {
      cb._st.sel = DOSEN_STATE === 'ok' ? '' : input.value.trim();
      cb.classList.remove('has-value');
      cb._st.active = -1;
      cbOpen(cb);
    });
    input.addEventListener('keydown', e => {
      const st = cb._st;
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        if (!cb.classList.contains('open')) return cbOpen(cb);
        const ok = st.items.map((it, i) => it.taken ? -1 : i).filter(i => i >= 0);
        if (!ok.length) return;
        let pos = ok.indexOf(st.active);
        pos = e.key === 'ArrowDown' ? (pos + 1) % ok.length : (pos <= 0 ? ok.length - 1 : pos - 1);
        st.active = ok[pos]; cbMarkActive(cb);
      } else if (e.key === 'Enter') {
        if (cb.classList.contains('open')) {
          e.preventDefault();
          const it = st.items[st.active >= 0 ? st.active : st.items.findIndex(x => !x.taken)];
          if (it && !it.taken) cbPick(cb, it.name);
        }
      } else if (e.key === 'Escape') { cbClose(cb); }
      else if (e.key === 'Tab') { cbClose(cb); }
    });
    panel.addEventListener('mousedown', e => e.preventDefault());   // jaga fokus input
    panel.addEventListener('click', e => {
      const el = e.target.closest('.cb-opt');
      if (!el) return;
      const it = cb._st.items[+el.dataset.i];
      if (it && !it.taken) cbPick(cb, it.name);
    });
    input.addEventListener('blur', () => cbClose(cb));
  }

  function cbOpen(cb) {
    document.querySelectorAll('.cb.open').forEach(o => { if (o !== cb) cbClose(o); });
    cb.classList.add('open');
    cb.querySelector('.cb-input').setAttribute('aria-expanded', 'true');
    cbRender(cb);
    const r = cb.getBoundingClientRect();
    cb.classList.toggle('up', window.innerHeight - r.bottom < 330 && r.top > window.innerHeight - r.bottom);
  }

  function cbClose(cb) {
    cb.classList.remove('open');
    const input = cb.querySelector('.cb-input');
    input.setAttribute('aria-expanded', 'false');
    if (DOSEN_STATE === 'ok') input.value = cb._st.sel;       // teks yang tidak dipilih dibuang
    cb.classList.toggle('has-value', !!cb._st.sel);
  }

  function cbPick(cb, name) {
    cb._st.sel = name;
    cb.querySelector('.cb-input').value = name;
    cbClose(cb);
    cb.querySelector('.cb-input').blur();
  }

  function cbRender(cb) {
    const st = cb._st, panel = cb.querySelector('.cb-panel');
    const input = cb.querySelector('.cb-input');
    if (DOSEN_STATE === 'loading') { panel.innerHTML = '<div class="cb-empty">Memuat daftar dosen…</div>'; return; }
    if (DOSEN_STATE === 'fail')    { panel.innerHTML = '<div class="cb-empty">Daftar dosen gagal dimuat.</div>'; return; }

    const q = st.sel === input.value ? '' : input.value.trim();     // baru dipilih -> tampilkan semua
    const toks = q.toLowerCase().split(/\s+/).filter(Boolean);
    const taken = new Set([...document.querySelectorAll('.cb')].filter(x => x !== cb && x._st.sel).map(x => x._st.sel));
    st.items = DAFTAR_DOSEN
      .filter(n => toks.every(t => n.toLowerCase().includes(t)))
      .map(n => ({ name: n, taken: taken.has(n) }));

    const meta = `<div class="cb-meta">${st.items.length} dosen${q ? ' ditemukan' : ' tersedia'}</div>`;
    if (!st.items.length) { panel.innerHTML = meta + '<div class="cb-empty">Tidak ada dosen yang cocok.<br>Coba kata kunci lain.</div>'; return; }

    const check = '<svg class="cb-check" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M4 10.5l4 4 8-9"/></svg>';
    panel.innerHTML = meta + '<div class="cb-list">' + st.items.map((it, i) => {
      const [nm, deg] = pecahNama(it.name);
      return `<div class="cb-opt${it.taken ? ' is-disabled' : ''}${it.name === st.sel ? ' is-selected' : ''}" role="option" data-i="${i}">
        <div class="cb-avatar">${esc(inisial(it.name))}</div>
        <div class="cb-txt"><div class="cb-name">${hl(nm, q)}</div>${deg ? `<div class="cb-deg">${hl(deg, q)}</div>` : ''}</div>
        ${it.taken ? '<span class="cb-tag">Dipilih</span>' : (it.name === st.sel ? check : '')}
      </div>`;
    }).join('') + '</div>';

    if (st.active < 0) st.active = Math.max(st.items.findIndex(x => x.name === st.sel), st.items.findIndex(x => !x.taken));
    cbMarkActive(cb);
  }

  function cbMarkActive(cb) {
    cb.querySelectorAll('.cb-opt').forEach(el => el.classList.toggle('is-active', +el.dataset.i === cb._st.active));
    const a = cb.querySelector('.cb-opt.is-active');
    if (a) a.scrollIntoView({ block: 'nearest' });
  }

  function init() {
    tambahDosen(); tambahDosen();
    tambahSkema();
    muatDosen();
  }

  function tambahDosen() {
    dosenCount++;
    const id = dosenCount;
    const row = document.createElement('div');
    row.className = 'dosen-row'; row.id = `dosenRow${id}`;
    row.innerHTML = `
      <div class="cb" id="cb${id}">
        <input type="text" class="cb-input" autocomplete="off" spellcheck="false" role="combobox" aria-expanded="false"
               aria-label="Nama Dosen ${id}" placeholder="Pilih nama dosen ${id}" />
        <svg class="cb-chev" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 8l5 5 5-5"/></svg>
        <div class="cb-panel" role="listbox"></div>
      </div>
      <input type="tel" inputmode="tel" placeholder="No. HP / WA (cth. +6281234567890)" oninput="filterHp(this)" />
      <button class="btn-remove" type="button" onclick="hapusRow('dosenRow${id}','dosenList','dosenRow',2)" title="Hapus">&#215;</button>`;
    document.getElementById('dosenList').appendChild(row);
    cbInit(row.querySelector('.cb'));
    updateRemoveButtons('dosenList', 'dosenRow', 2);
  }

  function filterHp(el) {
    // hanya angka dan tanda '+' (hanya di awal)
    let v = el.value.replace(/[^0-9+]/g, '');
    v = (v.startsWith('+') ? '+' : '') + v.replace(/\+/g, '');
    if (v !== el.value) el.value = v;
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
    return [...document.querySelectorAll('#dosenList .dosen-row')].map(row => ({
      nama: (row.querySelector('.cb')._st.sel || '').trim(),
      hp: (row.querySelector('input[type="tel"]').value || '').trim(),
    })).filter(d => d.nama);
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

  function setFormBusy(busy) {
    if (busy) document.querySelectorAll('.cb.open').forEach(cbClose);
    document.querySelectorAll('#formWrapper input, #formWrapper select, #formWrapper textarea, #formWrapper button')
      .forEach(el => { el.disabled = busy; });
    document.getElementById('formWrapper').setAttribute('aria-busy', busy ? 'true' : 'false');
    if (!busy) {   // pulihkan aturan tombol hapus (minimal baris)
      updateRemoveButtons('dosenList', 'dosenRow', 2);
      updateRemoveButtons('skemaList', 'skemaRow', 1);
    }
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
    if ([...document.querySelectorAll('#dosenList .cb')].some(c => c.querySelector('.cb-input').value.trim() && !c._st.sel)) errs.push('Nama dosen harus dipilih dari daftar dropdown.');
    if (payload.dosen.some(d => !d.hp)) errs.push('No. HP setiap dosen wajib diisi.');
    if (payload.dosen.some(d => d.hp && !/^\+?[0-9]+$/.test(d.hp))) errs.push('No. HP hanya boleh berisi angka dan tanda +.');
    skema.forEach((s, i) => {
      const n = i + 1;
      if (!s.judul) errs.push(`Judul skema ${n} wajib diisi.`);
      if (!s.file) errs.push(`File PDF skema ${n} wajib diunggah.`);
      else if (!/\.pdf$/i.test(s.file.name)) errs.push(`File skema ${n} harus berformat PDF.`);
      else if (s.file.size > MAX_PDF) errs.push(`File skema ${n} melebihi 4 MB.`);
    });
    if (errs.length) return showErrors(errs);

    setFormBusy(true); btn.classList.add('loading');
    const txt = btn.querySelector('.btn-text');

    try {
      txt.textContent = 'Memeriksa nama dosen...';
      const chk = await fetch('/check-dosen', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prodi: payload.prodi, dosen: payload.dosen }),
      });
      const chkJson = await chk.json();
      if (!chkJson.success) return showErrors(chkJson.errors || ['Nama dosen tidak valid.']);

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
      setFormBusy(false); btn.classList.remove('loading');
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


@app.route("/api/dosen")
def api_dosen():
    names = load_daftar_dosen()
    if not names:
        return jsonify({"success": False, "dosen": [], "errors": ["Daftar dosen gagal dimuat."]}), 502
    return jsonify({"success": True, "dosen": names})


@app.route("/check-dosen", methods=["POST"])
def check_dosen():
    try:
        data = request.get_json(force=True)
        prodi = (data.get("prodi") or "").strip()
        dosen, errs = kanonikkan_dosen(data.get("dosen", []))
        errs = errs or cek_dosen_duplikat(prodi, dosen)
        if errs:
            return jsonify({"success": False, "errors": errs}), 409
        return jsonify({"success": True})
    except Exception as e:
        logging.exception("Error di /check-dosen")
        return jsonify({"success": False, "errors": [str(e)]}), 500


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
        if any(not re.fullmatch(r"\+?[0-9]+", (d.get("hp") or "").strip()) for d in dosen_valid):
            errors.append("No. HP hanya boleh berisi angka dan tanda +.")
        if len(skema_valid) < 1: errors.append("Minimal 1 skema (judul + file PDF).")

        if errors:
            return jsonify({"success": False, "errors": errors}), 422

        dosen_valid, bad = kanonikkan_dosen(dosen_valid)
        if bad:
            return jsonify({"success": False, "errors": bad}), 422

        dup = cek_dosen_duplikat(prodi, dosen_valid)
        if dup:
            return jsonify({"success": False, "errors": dup}), 409

        row = [
            datetime.now(WITA).strftime("%Y-%m-%d %H:%M:%S"),
            prodi,
            format_dosen(dosen_valid),
            numbered([s["judul"].strip() for s in skema_valid]),
            ", ".join(s["url"].strip() for s in skema_valid),
        ]

        saved, errs = save_to_sheets(row)
        if not saved:
            save_to_csv(row)   # cadangan sementara (/tmp hilang saat redeploy)
            return jsonify({
                "success": False,
                "errors": ["Data gagal masuk ke Google Sheets."] + errs,
            }), 500

        return jsonify({"success": True})

    except Exception as e:
        logging.exception("Error di /submit")
        return jsonify({"success": False, "errors": [str(e)]}), 500


if __name__ == "__main__":
    app.run(debug=True)
