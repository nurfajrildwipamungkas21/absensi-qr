import streamlit as st
from datetime import datetime
from zoneinfo import ZoneInfo
import re
import io
from typing import Optional, Tuple, Dict, List
from collections import defaultdict
import difflib

import gspread
from google.oauth2.service_account import Credentials

import dropbox
from dropbox.sharing import RequestedVisibility, SharedLinkSettings
from dropbox.exceptions import ApiError, AuthError

import qrcode

# Pillow (biasanya sudah ada karena qrcode bergantung Pillow)
from PIL import Image, ImageOps


# =========================
# CONFIG
# =========================
st.set_page_config(
    page_title="JALA • Absensi QR",
    page_icon="✅",
    layout="centered",
    initial_sidebar_state="collapsed",
)

APP_CFG = st.secrets.get("app", {})
SHEET_NAME = APP_CFG.get("sheet_name", "Absensi_Karyawan")
WORKSHEET_NAME = APP_CFG.get("worksheet_name", "Log")
DROPBOX_ROOT = APP_CFG.get("dropbox_folder", "/Absensi_Selfie")
TZ_NAME = APP_CFG.get("timezone", "Asia/Jakarta")

QR_URL = APP_CFG.get("qr_url", "")
ENABLE_TOKEN = bool(APP_CFG.get("enable_token", False))
TOKEN_SECRET = str(APP_CFG.get("token", "")).strip()

# Kapasitas baris default agar format cukup jauh ke depan (hindari re-format tiap submit)
DEFAULT_SHEET_ROWS = int(APP_CFG.get("sheet_rows", 10000))

# Optimasi foto (server-side)
IMG_MAX_SIDE = int(APP_CFG.get("img_max_side", 1280))     # 1024-1600 aman untuk selfie
IMG_JPEG_QUALITY = int(APP_CFG.get("img_jpeg_quality", 78))  # 70-82 biasanya bagus

# Brand (bisa override via secrets)
BRAND_NAME = str(APP_CFG.get("brand_name", "JALA")).strip() or "JALA"
BRAND_TAGLINE = str(APP_CFG.get("brand_tagline", "Jala Tech")).strip() or "Jala Tech"
BRAND_PRIMARY = str(APP_CFG.get("brand_primary", "#0B66E4")).strip() or "#0B66E4"
BRAND_ACCENT = str(APP_CFG.get("brand_accent", "#46C2FF")).strip() or "#46C2FF"
BRAND_BG = str(APP_CFG.get("brand_bg", "#F5FAFF")).strip() or "#F5FAFF"

COL_TIMESTAMP = "Timestamp"
COL_NAMA = "Nama"
COL_HP = "No HP/WA"
COL_POSISI = "Posisi"
COL_LINK_SELFIE = "Bukti Selfie"     # tampil lebih professional
COL_DBX_PATH = "Dropbox Path"        # internal/admin

SHEET_COLUMNS = [COL_TIMESTAMP, COL_NAMA, COL_HP, COL_POSISI, COL_LINK_SELFIE, COL_DBX_PATH]


# =========================
# BRAND UI (CSS + HEADER)
# =========================
def inject_brand_css():
    st.markdown(
        f"""
<style>
:root {{
  --jala-primary: {BRAND_PRIMARY};
  --jala-accent: {BRAND_ACCENT};
  --jala-bg: {BRAND_BG};
  --jala-text: #0A2540;
  --jala-muted: #516579;
  --jala-border: rgba(11, 102, 228, 0.14);
  --jala-shadow: 0 10px 30px rgba(11, 102, 228, 0.18);
  --jala-card-shadow: 0 6px 18px rgba(11, 102, 228, 0.10);
}}

html, body, [data-testid="stAppViewContainer"] {{
  background: linear-gradient(180deg,
    rgba(70,194,255,0.16) 0%,
    rgba(245,250,255,1) 18%,
    #FFFFFF 100%) !important;
}}

[data-testid="stHeader"] {{
  background: rgba(255,255,255,0) !important;
}}

#MainMenu {{ visibility: hidden; }}
footer {{ visibility: hidden; }}

[data-testid="stAppViewContainer"] .main .block-container {{
  padding-top: 1.0rem;
  padding-bottom: 1.2rem;
  max-width: 720px;
}}

h1,h2,h3,h4 {{
  color: var(--jala-text) !important;
  letter-spacing: -0.01em;
}}

p, label, .stMarkdown, .stCaption {{
  color: var(--jala-text);
}}

.jala-topbar {{
  background: linear-gradient(135deg, var(--jala-accent) 0%, var(--jala-primary) 62%, #0B4CC7 100%);
  border-radius: 18px;
  padding: 16px 16px;
  box-shadow: var(--jala-shadow);
  margin: 0.25rem 0 1rem 0;
  overflow: hidden;
  position: relative;
}}

.jala-topbar:before {{
  content: "";
  position: absolute;
  top: -80px;
  right: -120px;
  width: 260px;
  height: 260px;
  background: rgba(255,255,255,0.16);
  border-radius: 999px;
  filter: blur(0px);
}}

.jala-brand {{
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  position: relative;
  z-index: 2;
}}

.jala-wordmark {{
  font-size: 26px;
  font-weight: 900;
  letter-spacing: 0.22em;
  color: #FFFFFF;
  line-height: 1.0;
}}

.jala-tagline {{
  margin-top: 6px;
  font-size: 13px;
  color: rgba(255,255,255,0.90);
}}

.jala-chip {{
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 7px 12px;
  border-radius: 999px;
  font-size: 12px;
  color: #FFFFFF;
  background: rgba(255,255,255,0.16);
  border: 1px solid rgba(255,255,255,0.22);
  white-space: nowrap;
}}

.jala-card {{
  background: rgba(255,255,255,0.92);
  border: 1px solid var(--jala-border);
  border-radius: 16px;
  padding: 14px 14px;
  box-shadow: var(--jala-card-shadow);
}}

.jala-muted {{
  color: var(--jala-muted);
  font-size: 13px;
}}

.jala-divider {{
  height: 1px;
  background: rgba(11,102,228,0.12);
  margin: 12px 0;
}}

div[data-testid="stButton"] button,
div[data-testid="stDownloadButton"] button {{
  border-radius: 14px !important;
  padding: 0.7rem 1rem !important;
}}

div[data-testid="stButton"] button[kind="primary"] {{
  background: linear-gradient(135deg, var(--jala-accent) 0%, var(--jala-primary) 82%) !important;
  border: 0 !important;
}}

div[data-testid="stButton"] button[kind="secondary"] {{
  border: 1px solid var(--jala-border) !important;
}}

div[data-testid="stForm"] {{
  border: 1px solid var(--jala-border);
  border-radius: 16px;
  padding: 14px;
  background: rgba(255,255,255,0.92);
  box-shadow: var(--jala-card-shadow);
}}

div[data-testid="stMetric"] {{
  background: rgba(255,255,255,0.92);
  border: 1px solid var(--jala-border);
  border-radius: 16px;
  padding: 10px;
  box-shadow: var(--jala-card-shadow);
}}

[data-testid="stInfo"], [data-testid="stWarning"], [data-testid="stError"], [data-testid="stSuccess"] {{
  border-radius: 14px;
}}

</style>
        """,
        unsafe_allow_html=True,
    )


def render_header(chip_text: str, subtitle: str):
    st.markdown(
        f"""
<div class="jala-topbar">
  <div class="jala-brand">
    <div>
      <div class="jala-wordmark">{BRAND_NAME}</div>
      <div class="jala-tagline">{subtitle}</div>
    </div>
    <div class="jala-chip">{chip_text}</div>
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )


inject_brand_css()


# =========================
# HELPERS
# =========================
def get_mode() -> str:
    # kompatibel streamlit baru & lama
    try:
        return str(st.query_params.get("mode", "")).strip().lower()
    except Exception:
        qp = st.experimental_get_query_params()
        return (qp.get("mode", [""])[0] or "").strip().lower()


def get_token_from_url() -> str:
    try:
        return str(st.query_params.get("token", "")).strip()
    except Exception:
        qp = st.experimental_get_query_params()
        return (qp.get("token", [""])[0] or "").strip()


def sanitize_name(text: str) -> str:
    text = str(text).strip()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^A-Za-z0-9 _.-]", "", text)
    return text.strip()


def sanitize_phone(text: str) -> str:
    text = str(text).strip()
    if text.startswith("+"):
        return "+" + re.sub(r"\D", "", text[1:])
    return re.sub(r"\D", "", text)


def now_local():
    return datetime.now(tz=ZoneInfo(TZ_NAME))


@st.cache_data(show_spinner=False)
def build_qr_png(url: str) -> bytes:
    # QR cukup tajam tapi tidak berlebihan agar cepat dimuat
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def make_hyperlink(url: str, label: str = "Bukti Foto") -> str:
    """Supaya kolom link rapi di GSheet/Excel."""
    if not url or url == "-":
        return "-"
    safe = url.replace('"', '""')  # escape double quote untuk formula
    return f'=HYPERLINK("{safe}", "{label}")'


def detect_ext_and_mime(mime: str) -> str:
    mime = (mime or "").lower()
    if "png" in mime:
        return ".png"
    return ".jpg"


def get_selfie_bytes(selfie_cam, selfie_upload) -> Tuple[Optional[bytes], str]:
    """Return (bytes, ext)."""
    if selfie_cam is not None:
        mime = getattr(selfie_cam, "type", "") or ""
        return selfie_cam.getvalue(), detect_ext_and_mime(mime)
    if selfie_upload is not None:
        mime = getattr(selfie_upload, "type", "") or ""
        return selfie_upload.getvalue(), detect_ext_and_mime(mime)
    return None, ".jpg"


def optimize_image_bytes(img_bytes: bytes, ext: str) -> Tuple[bytes, str]:
    """
    Optimasi server-side:
    - perbaiki orientasi EXIF
    - resize max side (default 1280)
    - kompres ke JPEG berkualitas baik
    Tujuan: lebih cepat proses + lebih hemat storage/bandwidth Dropbox.
    """
    try:
        img = Image.open(io.BytesIO(img_bytes))
        img = ImageOps.exif_transpose(img)

        # Convert ke RGB (wajib untuk JPEG)
        if img.mode not in ("RGB", "L"):
            # kalau ada alpha, campur dengan background putih
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode in ("RGBA", "LA"):
                bg.paste(img, mask=img.split()[-1])
            else:
                bg.paste(img)
            img = bg
        else:
            img = img.convert("RGB")

        w, h = img.size
        max_side = max(w, h)
        if max_side > IMG_MAX_SIDE:
            scale = IMG_MAX_SIDE / float(max_side)
            new_size = (int(w * scale), int(h * scale))
            img = img.resize(new_size, Image.LANCZOS)

        out = io.BytesIO()
        img.save(
            out,
            format="JPEG",
            quality=IMG_JPEG_QUALITY,
            optimize=True,
            progressive=True,
        )
        return out.getvalue(), ".jpg"
    except Exception:
        # fallback: pakai bytes original
        return img_bytes, ext


# =========================
# SHEET FORMATTER (lebih hemat request)
# =========================
def auto_format_absensi_sheet(ws):
    """Format Google Sheet Absensi agar rapi & profesional (tanpa baca all_values)."""
    try:
        sheet_id = ws.id
        row_count = ws.row_count  # gunakan kapasitas sheet (diset besar saat create)

        # Lebar kolom A-F
        # A Timestamp, B Nama, C No HP/WA, D Posisi, E Bukti Selfie, F Dropbox Path
        col_widths = [170, 180, 150, 180, 140, 340]

        requests = []

        # 1) Set lebar kolom
        for i, w in enumerate(col_widths):
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i + 1
                    },
                    "properties": {"pixelSize": w},
                    "fields": "pixelSize"
                }
            })

        # 2) Header styling (row 1)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {"userEnteredFormat": {
                    "textFormat": {"bold": True},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "backgroundColor": {"red": 0.93, "green": 0.95, "blue": 0.99},
                    "wrapStrategy": "WRAP"
                }},
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor,wrapStrategy)"
            }
        })

        # 3) Freeze header
        requests.append({
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount"
            }
        })

        # 4) Body default format
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": row_count},
                "cell": {"userEnteredFormat": {
                    "verticalAlignment": "MIDDLE",
                    "wrapStrategy": "CLIP"
                }},
                "fields": "userEnteredFormat(verticalAlignment,wrapStrategy)"
            }
        })

        # 5) Center: Timestamp (A) & No HP (C)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": 0, "endColumnIndex": 1},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat(horizontalAlignment)"
            }
        })
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": 2, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat(horizontalAlignment)"
            }
        })

        # 6) Wrap untuk Dropbox Path (F)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": 5, "endColumnIndex": 6},
                "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP"}},
                "fields": "userEnteredFormat(wrapStrategy)"
            }
        })

        if requests:
            ws.spreadsheet.batch_update({"requests": requests})

    except Exception as e:
        # jangan bikin app crash kalau format gagal
        print(f"Format Absensi Error: {e}")


@st.cache_resource
def connect_gsheet():
    if "gcp_service_account" not in st.secrets:
        raise RuntimeError("GSheet secrets tidak ditemukan: gcp_service_account")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open(SHEET_NAME)


def get_or_create_ws(spreadsheet):
    try:
        ws = spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=WORKSHEET_NAME,
            rows=DEFAULT_SHEET_ROWS,
            cols=len(SHEET_COLUMNS),
        )
        ws.append_row(SHEET_COLUMNS, value_input_option="USER_ENTERED")
        auto_format_absensi_sheet(ws)
        return ws

    # Pastikan kapasitas baris cukup (sekali, bukan tiap submit)
    if ws.row_count < DEFAULT_SHEET_ROWS:
        ws.resize(rows=DEFAULT_SHEET_ROWS)

    header = ws.row_values(1)
    if header != SHEET_COLUMNS:
        ws.resize(cols=max(ws.col_count, len(SHEET_COLUMNS)))
        ws.update("A1", [SHEET_COLUMNS], value_input_option="USER_ENTERED")
        auto_format_absensi_sheet(ws)

    return ws


@st.cache_resource
def connect_dropbox():
    if "dropbox" not in st.secrets or "access_token" not in st.secrets["dropbox"]:
        raise RuntimeError("Dropbox secrets tidak ditemukan: dropbox.access_token")

    dbx = dropbox.Dropbox(st.secrets["dropbox"]["access_token"])
    dbx.users_get_current_account()
    return dbx


def upload_selfie_to_dropbox(
    dbx,
    img_bytes: bytes,
    nama: str,
    ts_file: str,
    ext: str
) -> Tuple[str, str]:
    """
    Return (shared_link_raw, dropbox_path)
    """
    clean_name = sanitize_name(nama).replace(" ", "_") or "Unknown"
    filename = f"{ts_file}_selfie{ext}"
    path = f"{DROPBOX_ROOT}/{clean_name}/{filename}"

    dbx.files_upload(img_bytes, path, mode=dropbox.files.WriteMode.add)

    settings = SharedLinkSettings(requested_visibility=RequestedVisibility.public)
    url = "-"
    try:
        link = dbx.sharing_create_shared_link_with_settings(path, settings=settings)
        url = link.url
    except ApiError as e:
        try:
            if e.error.is_shared_link_already_exists():
                links = dbx.sharing_list_shared_links(path, direct_only=True).links
                if links:
                    url = links[0].url
        except Exception:
            url = "-"

    url_raw = url.replace("?dl=0", "?raw=1") if url and url != "-" else "-"
    return url_raw, path


# =========================
# REKAP (PINTAR) - POSISI & HADIR
# =========================
def normalize_posisi(text: str) -> str:
    t = str(text or "").strip().lower()
    t = t.replace("&", " dan ")
    t = re.sub(r"[/,_\-\.]+", " ", t)
    t = re.sub(r"[^a-z0-9\s]+", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


POSISI_ALIASES: Dict[str, str] = {
    "spv": "supervisor",
    "sup": "supervisor",
    "super visor": "supervisor",
    "supervisior": "supervisor",
    "admin": "administrasi",
    "adm": "administrasi",
    "kry": "karyawan",
    "karyawan": "karyawan",
    "staf": "staff",
    "staff": "staff",
    "teknisi": "teknisi",
    "technician": "teknisi",
    "driver": "driver",
    "drv": "driver",
    "security": "security",
    "satpam": "security",
}


def smart_canonical_posisi(raw_pos: str, known_canon: List[str]) -> str:
    p = normalize_posisi(raw_pos)
    if not p:
        return ""

    if p in POSISI_ALIASES:
        p = POSISI_ALIASES[p]

    if known_canon:
        best = difflib.get_close_matches(p, known_canon, n=1, cutoff=0.88)
        if best:
            return best[0]

    return p


def display_posisi(canon: str) -> str:
    if not canon:
        return "-"
    return " ".join(w.capitalize() for w in canon.split())


def parse_date_prefix(ts: str) -> str:
    s = str(ts or "").strip()
    if not s:
        return ""
    try:
        dt = datetime.strptime(s, "%d-%m-%Y %H:%M:%S")
        return dt.strftime("%d-%m-%Y")
    except Exception:
        return s[:10]


def _group_contiguous_rows(rows: List[int]) -> List[Tuple[int, int]]:
    if not rows:
        return []
    rows = sorted(rows)
    ranges = []
    start = prev = rows[0]
    for r in rows[1:]:
        if r == prev + 1:
            prev = r
        else:
            ranges.append((start, prev))
            start = prev = r
    ranges.append((start, prev))
    return ranges


@st.cache_data(ttl=30, show_spinner=False)
def get_rekap_today() -> Dict:
    """
    Rekap hari ini (lebih hemat data):
    - Ambil kolom Timestamp (A) dulu
    - Ambil hanya baris yang tanggalnya hari ini untuk A:D
    """
    sh = connect_gsheet()
    ws = get_or_create_ws(sh)

    today_str = now_local().strftime("%d-%m-%Y")

    # Ambil kolom A (Timestamp) saja: jauh lebih ringan daripada A:D semua baris
    ts_col = ws.col_values(1)  # termasuk header
    if not ts_col or len(ts_col) < 2:
        return {
            "today": today_str,
            "total": 0,
            "dup_removed": 0,
            "by_pos": [],
            "all_people": [],
        }

    # Cari row index yang match hari ini
    match_rows = []
    for idx, ts in enumerate(ts_col[1:], start=2):
        if parse_date_prefix(ts) == today_str:
            match_rows.append(idx)

    if not match_rows:
        return {
            "today": today_str,
            "total": 0,
            "dup_removed": 0,
            "by_pos": [],
            "all_people": [],
        }

    # Fetch hanya range yang diperlukan (minim request: gabungkan range yang kontigu)
    ranges = _group_contiguous_rows(match_rows)
    data = []
    for a, b in ranges:
        chunk = ws.get(f"A{a}:D{b}")
        if chunk:
            data.extend(chunk)

    # Dedup: kunci utama No HP (lebih unik), fallback Nama.
    seen_keys = set()
    dup_removed = 0

    people_by_pos = defaultdict(list)
    all_people = []
    known_canon = []

    for r in data:
        ts = (r[0] if len(r) > 0 else "") or ""
        nama = (r[1] if len(r) > 1 else "") or ""
        hp = (r[2] if len(r) > 2 else "") or ""
        pos = (r[3] if len(r) > 3 else "") or ""

        if parse_date_prefix(ts) != today_str:
            continue

        nama_clean = sanitize_name(nama)
        hp_clean = sanitize_phone(hp)
        key = hp_clean if hp_clean else nama_clean.lower().strip()

        if not key:
            continue

        if key in seen_keys:
            dup_removed += 1
            continue

        seen_keys.add(key)

        pos_canon = smart_canonical_posisi(pos, known_canon)
        if pos_canon and pos_canon not in known_canon:
            known_canon.append(pos_canon)

        who = nama_clean if nama_clean else (hp_clean if hp_clean else "Tanpa Nama")
        who_display = f"{who} ({hp_clean})" if hp_clean and who else who

        all_people.append({
            "Nama": who,
            "No HP/WA": hp_clean or "-",
            "Posisi": display_posisi(pos_canon) if pos_canon else "-",
            "Timestamp": ts,
        })

        people_by_pos[pos_canon if pos_canon else "(tanpa posisi)"].append(who_display)

    by_pos = []
    for canon, people in people_by_pos.items():
        by_pos.append({
            "Posisi": display_posisi(canon) if canon != "(tanpa posisi)" else "Tanpa Posisi",
            "Jumlah": len(people),
            "Yang Hadir": ", ".join(people),
        })

    by_pos.sort(key=lambda x: (-x["Jumlah"], x["Posisi"].lower()))

    return {
        "today": today_str,
        "total": len(seen_keys),
        "dup_removed": dup_removed,
        "by_pos": by_pos,
        "all_people": all_people,
    }


# =========================
# SESSION DEFAULTS
# =========================
if "saving" not in st.session_state:
    st.session_state.saving = False
if "submitted_once" not in st.session_state:
    st.session_state.submitted_once = False
if "selfie_method" not in st.session_state:
    st.session_state.selfie_method = "Upload"  # default hemat resource


# =========================
# UI
# =========================
mode = get_mode()

# ===== PAGE: QR / ADMIN
if mode != "absen":
    render_header("QR Absensi", f"{BRAND_TAGLINE} • Scan → Isi Form")

    st.markdown(
        """
<div class="jala-card">
  <div style="font-weight:700; font-size:16px; margin-bottom:6px;">QR Code Absensi</div>
  <div class="jala-muted">
    Cetak/Tempel QR ini di lokasi masuk. Karyawan scan QR → langsung ke form absen.
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")

    if not QR_URL:
        st.warning("QR URL belum diset. Isi `app.qr_url` di secrets.")
        st.code("Contoh: https://YOUR-APP.streamlit.app/?mode=absen", language="text")
        st.stop()

    if ENABLE_TOKEN and TOKEN_SECRET:
        if "token=" not in QR_URL:
            sep = "&" if "?" in QR_URL else "?"
            qr_url_effective = f"{QR_URL}{sep}token={TOKEN_SECRET}"
        else:
            qr_url_effective = QR_URL
    else:
        qr_url_effective = QR_URL

    qr_png = build_qr_png(qr_url_effective)

    st.markdown('<div class="jala-card">', unsafe_allow_html=True)
    st.image(qr_png, caption="QR Absensi", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        st.link_button("🔗 Tes Link Absensi", qr_url_effective, use_container_width=True)
    with c2:
        st.download_button(
            "⬇️ Download QR",
            data=qr_png,
            file_name="qr_absensi_jala.png",
            mime="image/png",
            use_container_width=True
        )

    with st.expander("ℹ️ Tips Penggunaan"):
        st.write(
            "- Pastikan URL aplikasi **HTTPS**.\n"
            "- Untuk HP jadul: gunakan **Upload foto** (kamera browser kadang tidak stabil).\n"
            "- Jika pakai token, QR mengandung `token=...` agar tidak sembarang orang submit.\n"
            "- Untuk koneksi lambat: gunakan foto dari kamera (biasanya lebih kecil)."
        )

    st.markdown(
        f"""
<div style="text-align:center; margin-top: 10px;" class="jala-muted">
  © {BRAND_TAGLINE} • Absensi QR
</div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()


# ===== PAGE: ABSEN (dibuka dari scan QR)
dt = now_local()
ts_display = dt.strftime("%d-%m-%Y %H:%M:%S")
ts_file = dt.strftime("%Y-%m-%d_%H-%M-%S")

render_header("Form Absensi", f"{BRAND_TAGLINE} • {ts_display} ({TZ_NAME})")

if ENABLE_TOKEN and TOKEN_SECRET:
    incoming_token = get_token_from_url()
    if incoming_token != TOKEN_SECRET:
        st.error("Akses tidak valid. Silakan scan QR resmi dari kantor.")
        st.stop()

st.markdown(
    """
<div class="jala-card">
  <div style="font-weight:700; font-size:16px; margin-bottom:6px;">Petunjuk Singkat</div>
  <div class="jala-muted">
    Isi data karyawan lalu unggah/ambil selfie kehadiran.
    Jika ada izin kamera, pilih <b>Allow / Izinkan</b>.
    Untuk HP tertentu, pilih metode <b>Upload</b>.
  </div>
</div>
    """,
    unsafe_allow_html=True,
)

st.write("")

with st.form("form_absen", clear_on_submit=False):
    st.subheader("1) Data Karyawan")
    nama = st.text_input("Nama Lengkap", placeholder="Contoh: Andi Saputra")
    no_hp = st.text_input("No HP/WA", placeholder="Contoh: 08xxxxxxxxxx atau +628xxxxxxxxxx")
    posisi = st.text_input("Posisi / Jabatan", placeholder="Contoh: Driver / Teknisi / Supervisor")

    st.markdown('<div class="jala-divider"></div>', unsafe_allow_html=True)

    st.subheader("2) Selfie Kehadiran")

    # Pilihan metode agar kamera tidak berat (dan tidak auto load)
    method = st.radio(
        "Pilih metode selfie",
        options=["Upload (lebih stabil)", "Kamera (jika HP mendukung)"],
        index=0 if st.session_state.selfie_method == "Upload" else 1,
        horizontal=False,
    )
    st.session_state.selfie_method = "Upload" if method.startswith("Upload") else "Kamera"

    selfie_cam = None
    selfie_upload = None

    if st.session_state.selfie_method == "Kamera":
        st.caption("Jika kamera blank/lemot, kembali pilih metode Upload.")
        selfie_cam = st.camera_input("Ambil selfie")
    else:
        st.caption("Disarankan pilih foto yang tidak terlalu besar. Sistem akan mengoptimalkan foto secara otomatis.")
        selfie_upload = st.file_uploader("Upload foto selfie", type=["jpg", "jpeg", "png"])

    st.markdown('<div class="jala-divider"></div>', unsafe_allow_html=True)

    submit = st.form_submit_button(
        "✅ Submit Absensi",
        disabled=st.session_state.saving or st.session_state.submitted_once,
        use_container_width=True,
        type="primary",
    )

# ===== SUBMIT LOGIC
if submit:
    if st.session_state.submitted_once:
        st.warning("Absensi sudah tersimpan. Jika ingin absen lagi, refresh halaman.")
        st.stop()

    nama_clean = sanitize_name(nama)
    hp_clean = sanitize_phone(no_hp)
    posisi_final = str(posisi).strip()

    img_bytes, ext = get_selfie_bytes(selfie_cam, selfie_upload)

    errors = []
    if not nama_clean:
        errors.append("• Nama wajib diisi.")
    if not hp_clean or len(hp_clean.replace("+", "")) < 8:
        errors.append("• No HP/WA wajib diisi (minimal 8 digit).")
    if not posisi_final:
        errors.append("• Posisi wajib diisi.")
    if img_bytes is None:
        errors.append("• Selfie wajib (kamera atau upload).")

    if errors:
        st.error("Mohon lengkapi dulu:\n\n" + "\n".join(errors))
        st.stop()

    st.session_state.saving = True
    try:
        with st.spinner("Menyimpan absensi..."):
            # Optimasi foto (lebih ringan untuk proses & Dropbox)
            img_bytes_opt, ext_opt = optimize_image_bytes(img_bytes, ext)

            sh = connect_gsheet()
            ws = get_or_create_ws(sh)
            dbx = connect_dropbox()

            link_selfie, dbx_path = upload_selfie_to_dropbox(
                dbx, img_bytes_opt, nama_clean, ts_file, ext_opt
            )

            link_cell = make_hyperlink(link_selfie, "Bukti Foto")

            ws.append_row(
                [ts_display, nama_clean, hp_clean, posisi_final, link_cell, dbx_path],
                value_input_option="USER_ENTERED"
            )

        # setelah submit, rekap perlu refresh
        get_rekap_today.clear()

        st.session_state.submitted_once = True
        st.success("Absensi berhasil tersimpan. Terima kasih ✅")
        st.balloons()

        if st.button("↩️ Isi ulang (reset form)", use_container_width=True):
            st.session_state.saving = False
            st.session_state.submitted_once = False
            st.session_state.selfie_method = "Upload"
            st.rerun()

    except AuthError:
        st.error("Dropbox token tidak valid. Hubungi admin.")
    except Exception as e:
        st.error("Gagal menyimpan absensi.")
        with st.expander("Detail error (untuk admin)"):
            st.code(str(e))
    finally:
        st.session_state.saving = False


# =========================
# UI: REKAP KEHADIRAN
# =========================
st.write("")
st.subheader("📊 Rekap Kehadiran (Hari ini)")

try:
    rekap = get_rekap_today()

    top1, top2 = st.columns([1, 1])
    with top1:
        st.metric("Total hadir", rekap["total"])
    with top2:
        if st.button("🔄 Refresh rekap", use_container_width=True):
            get_rekap_today.clear()
            st.rerun()

    st.caption(f"Tanggal: **{rekap['today']}**")

    if rekap["dup_removed"] > 0:
        st.info(
            f"Catatan: terdeteksi **{rekap['dup_removed']}** entri duplikat "
            f"(No HP/Nama sama) dan tidak dihitung agar rekap akurat."
        )

    if rekap["total"] == 0:
        st.warning("Belum ada absensi untuk hari ini.")
    else:
        st.markdown(
            """
<div class="jala-card">
  <div style="font-weight:700; margin-bottom:8px;">Klasifikasi jumlah hadir per posisi</div>
</div>
            """,
            unsafe_allow_html=True,
        )
        st.dataframe(rekap["by_pos"], use_container_width=True, hide_index=True)

        with st.expander("👥 Lihat siapa saja yang sudah datang (detail)"):
            st.dataframe(rekap["all_people"], use_container_width=True, hide_index=True)

except Exception as e:
    st.warning("Rekap kehadiran belum bisa ditampilkan (cek koneksi GSheet).")
    with st.expander("Detail error (untuk admin)"):
        st.code(str(e))

st.markdown(
    f"""
<div style="text-align:center; margin-top: 12px;" class="jala-muted">
  © {BRAND_TAGLINE} • Absensi QR
</div>
    """,
    unsafe_allow_html=True,
)
