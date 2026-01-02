import streamlit as st
from datetime import datetime
from zoneinfo import ZoneInfo
import re
import io
import csv
import difflib
from collections import Counter
from typing import Optional, Tuple, List, Dict, Any

import gspread
from google.oauth2.service_account import Credentials

import dropbox
from dropbox.sharing import RequestedVisibility, SharedLinkSettings
from dropbox.exceptions import ApiError, AuthError

import qrcode

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Absensi QR", page_icon="✅", layout="centered")

APP_CFG = st.secrets.get("app", {})
SHEET_NAME = APP_CFG.get("sheet_name", "Absensi_Karyawan")
WORKSHEET_NAME = APP_CFG.get("worksheet_name", "Log")
DROPBOX_ROOT = APP_CFG.get("dropbox_folder", "/Absensi_Selfie")
TZ_NAME = APP_CFG.get("timezone", "Asia/Jakarta")

QR_URL = APP_CFG.get("qr_url", "")
ENABLE_TOKEN = bool(APP_CFG.get("enable_token", False))
TOKEN_SECRET = str(APP_CFG.get("token", "")).strip()

# Kolom lama (tetap dipertahankan biar kompatibel)
COL_TIMESTAMP = "Timestamp"
COL_NAMA = "Nama"
COL_HP = "No HP/WA"
COL_POSISI = "Posisi"
COL_LINK_SELFIE = "Bukti Selfie"     # tampil lebih professional
COL_DBX_PATH = "Dropbox Path"        # internal/admin

# Kolom tambahan untuk akurasi rekap & download (tidak mengganggu UI sheet)
COL_POSISI_NORM = "Posisi (Normalized)"   # internal (untuk rekap pintar)
COL_SELFIE_RAW = "Selfie URL Raw"         # internal (biar download tidak meleset)

# Penting: 6 kolom awal SAMA seperti sebelumnya, yang baru ditambahkan di belakang
SHEET_COLUMNS = [
    COL_TIMESTAMP, COL_NAMA, COL_HP, COL_POSISI, COL_LINK_SELFIE, COL_DBX_PATH,
    COL_POSISI_NORM, COL_SELFIE_RAW
]


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


def build_qr_png(url: str) -> bytes:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def make_hyperlink(url: str, label: str = "Bukti Foto") -> str:
    """Supaya kolom link rapi di GSheet/Excel."""
    if not url or url == "-":
        return "-"
    safe = url.replace('"', '""')  # escape double quote untuk formula
    return f'=HYPERLINK("{safe}", "{label}")'


# ---- Normalisasi posisi (pintar tapi aman)
def _default_pos_aliases() -> Dict[str, str]:
    """
    Alias default (bisa ditambah via secrets: app.position_aliases).
    Tips: di secrets.toml bisa buat:
    [app.position_aliases]
    spv="supervisor"
    sup="supervisor"
    security="satpam"
    """
    return {
        "spv": "supervisor",
        "sup": "supervisor",
        "super visor": "supervisor",
        "supervisor": "supervisor",
        "leader": "leader",
        "ketua": "leader",
        "admin": "admin",
        "operator": "operator",
        "ops": "operator",
        "teknisi": "teknisi",
        "technician": "teknisi",
        "driver": "driver",
        "supir": "driver",
        "satpam": "satpam",
        "security": "satpam",
        "karyawan": "karyawan",
        "pegawai": "karyawan",
        "staff": "karyawan",
        "staf": "karyawan",
        "warehouse": "gudang",
        "gudang": "gudang",
    }


def get_pos_aliases() -> Dict[str, str]:
    user_alias = APP_CFG.get("position_aliases", {}) or {}
    # normalisasi key/value user
    merged = _default_pos_aliases()
    for k, v in dict(user_alias).items():
        kk = normalize_posisi(str(k))
        vv = normalize_posisi(str(v))
        if kk and vv:
            merged[kk] = vv
    return merged


def normalize_posisi(text: str) -> str:
    """
    Normalisasi stabil untuk menghindari kategori terpecah:
    - lowercase
    - ganti pemisah (/,-,_) jadi spasi
    - hapus karakter aneh
    - rapikan spasi
    """
    t = str(text or "").strip().lower()
    if not t:
        return ""
    t = t.replace("&", " dan ")
    t = re.sub(r"[\/\-\_\.\|]+", " ", t)
    t = re.sub(r"[^a-z0-9\s]+", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def canonicalize_posisi(raw_pos: str, known_canon: Optional[List[str]] = None) -> str:
    """
    "Pintar" tapi aman:
    1) normalisasi dasar
    2) alias mapping (spv->supervisor, dll)
    3) fuzzy match SANGAT ketat (hindari salah merge)
    """
    base = normalize_posisi(raw_pos)
    if not base:
        return ""

    aliases = get_pos_aliases()
    if base in aliases:
        return aliases[base]

    # coba alias pada token gabungan (mis. "spv lapangan" => "supervisor lapangan" jika kamu set di secrets)
    # tapi defaultnya kita biarkan base apa adanya agar tidak salah
    candidates = []
    if known_canon:
        candidates.extend([c for c in known_canon if c])

    # tambahkan nilai alias (biar "super visor" bisa nyambung ke "supervisor")
    candidates.extend(list(set(aliases.values())))

    candidates = list(dict.fromkeys(candidates))  # unique preserve order
    if not candidates:
        return base

    # fuzzy match ketat: hanya untuk typo kecil (mis. "supervissor" -> "supervisor")
    best = difflib.get_close_matches(base, candidates, n=1, cutoff=0.92)
    if best:
        return best[0]

    return base


def display_posisi(norm: str) -> str:
    if not norm:
        return "-"
    return " ".join([w.capitalize() for w in norm.split(" ")])


def ts_to_datekey(ts: str) -> str:
    """
    Ambil "dd-mm-YYYY" dari Timestamp.
    Aman untuk format utama: "dd-mm-YYYY HH:MM:SS".
    """
    s = str(ts or "").strip()
    if len(s) >= 10 and s[2:3] == "-" and s[5:6] == "-":
        return s[:10]
    # fallback parsing
    try:
        dt = datetime.strptime(s, "%d-%m-%Y %H:%M:%S")
        return dt.strftime("%d-%m-%Y")
    except Exception:
        return ""


def to_csv_bytes(rows: List[Dict[str, Any]], columns: List[str]) -> bytes:
    """
    CSV UTF-8 with BOM (biar Excel enak).
    """
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(columns)
    for r in rows:
        writer.writerow([r.get(c, "") for c in columns])
    text = output.getvalue()
    return ("\ufeff" + text).encode("utf-8")


def auto_format_absensi_sheet(ws):
    """Format Google Sheet Absensi agar rapi & profesional."""
    try:
        sheet_id = ws.id
        all_values = ws.get_all_values()
        row_count = max(len(all_values), ws.row_count)

        # Lebar kolom A-H (disesuaikan dengan kolom tambahan)
        # A Timestamp, B Nama, C No HP/WA, D Posisi, E Bukti Selfie, F Dropbox Path, G Posisi Norm, H Selfie Raw
        col_widths = [170, 180, 150, 180, 140, 340, 190, 320]

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
                    "backgroundColor": {"red": 0.93, "green": 0.93, "blue": 0.93},
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

        # 6) Wrap untuk Dropbox Path (F) dan Selfie Raw (H)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": 5, "endColumnIndex": 6},
                "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP"}},
                "fields": "userEnteredFormat(wrapStrategy)"
            }
        })
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": 7, "endColumnIndex": 8},
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
        ws = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=5000, cols=len(SHEET_COLUMNS))
        ws.append_row(SHEET_COLUMNS, value_input_option="USER_ENTERED")
        auto_format_absensi_sheet(ws)
        return ws

    header = ws.row_values(1)
    if header != SHEET_COLUMNS:
        # upgrade header dengan aman
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


def detect_ext_and_mime(mime: str) -> str:
    mime = (mime or "").lower()
    if "png" in mime:
        return ".png"
    return ".jpg"


def get_selfie_bytes(selfie_cam, selfie_upload) -> Tuple[Optional[bytes], str]:
    """
    Return (bytes, ext).
    """
    if selfie_cam is not None:
        mime = getattr(selfie_cam, "type", "") or ""
        return selfie_cam.getvalue(), detect_ext_and_mime(mime)

    if selfie_upload is not None:
        mime = getattr(selfie_upload, "type", "") or ""
        return selfie_upload.getvalue(), detect_ext_and_mime(mime)

    return None, ".jpg"


def already_checked_in_today(ws, hp_clean: str, today_key: str) -> Tuple[bool, str]:
    """
    Cegah double absen (biar rekap & download tidak meleset).
    Cek berdasarkan No HP + tanggal.
    Return (bool, timestamp_terakhir)
    """
    hp_clean = (hp_clean or "").strip()
    if not hp_clean:
        return False, ""

    # Ambil kolom timestamp dan hp (lebih ringan daripada get_all_values)
    ts_list = ws.col_values(1)[1:]   # A, skip header
    hp_list = ws.col_values(3)[1:]   # C, skip header

    # scan dari bawah (yang terbaru) biar cepat ketemu
    for ts, hp in zip(reversed(ts_list), reversed(hp_list)):
        if ts_to_datekey(ts) != today_key:
            # begitu sudah lewat tanggal hari ini, boleh break (karena urutan append biasanya naik)
            # tapi untuk aman, kita tidak break keras (kadang sheet bisa di-sort).
            continue
        if sanitize_phone(hp) == hp_clean:
            return True, str(ts)
    return False, ""


@st.cache_data(ttl=30)
def get_today_data_and_rekap() -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], int, str]:
    """
    Return:
      - daftar hadir hari ini (rows)
      - rekap posisi hari ini (rows)
      - total hadir
      - today_key (dd-mm-YYYY)
    """
    sh = connect_gsheet()
    ws = get_or_create_ws(sh)

    today_key = now_local().strftime("%d-%m-%Y")
    records = ws.get_all_records(default_blank="")  # pakai header row

    # siapkan known canon dari data yang sudah ada hari ini (untuk fuzzy match aman)
    known = []
    for r in records:
        if ts_to_datekey(r.get(COL_TIMESTAMP, "")) != today_key:
            continue
        pnorm = normalize_posisi(r.get(COL_POSISI_NORM, "")) or normalize_posisi(r.get(COL_POSISI, ""))
        if pnorm:
            known.append(pnorm)
    known = list(dict.fromkeys(known))

    hadir_today: List[Dict[str, Any]] = []
    counter = Counter()

    for r in records:
        ts = r.get(COL_TIMESTAMP, "")
        if ts_to_datekey(ts) != today_key:
            continue

        nama = str(r.get(COL_NAMA, "")).strip()
        hp = str(r.get(COL_HP, "")).strip()
        posisi_raw = str(r.get(COL_POSISI, "")).strip()

        # ambil normalized jika sudah tersimpan, kalau kosong => hitung dari posisi raw
        posisi_norm_saved = str(r.get(COL_POSISI_NORM, "")).strip()
        posisi_norm = normalize_posisi(posisi_norm_saved) if posisi_norm_saved else ""
        if not posisi_norm:
            posisi_norm = canonicalize_posisi(posisi_raw, known_canon=known)

        # selfie url raw untuk download (kalau kosong, tetap simpan apa adanya)
        selfie_raw = str(r.get(COL_SELFIE_RAW, "")).strip()
        # Bukti Selfie di sheet mungkin formula hyperlink; untuk UI kita cukup tampilkan "ada/tidak"
        bukti = str(r.get(COL_LINK_SELFIE, "")).strip()

        posisi_disp = display_posisi(posisi_norm)
        counter[posisi_disp] += 1

        hadir_today.append({
            "Timestamp": ts,
            "Nama": nama,
            "No HP/WA": hp,
            "Posisi": posisi_disp,
            "Selfie URL": selfie_raw if selfie_raw else "",  # untuk download biar akurat
            "Bukti (Sheet)": bukti,
        })

    total = sum(counter.values())
    rekap_rows = [{"Posisi": k, "Jumlah": v} for k, v in sorted(counter.items(), key=lambda x: (-x[1], x[0]))]
    return hadir_today, rekap_rows, total, today_key


# =========================
# SESSION DEFAULTS
# =========================
if "open_cam" not in st.session_state:
    st.session_state.open_cam = False
if "saving" not in st.session_state:
    st.session_state.saving = False
if "submitted_once" not in st.session_state:
    st.session_state.submitted_once = False


# =========================
# UI
# =========================
mode = get_mode()

# ===== PAGE: QR / ADMIN
if mode != "absen":
    st.title("✅ QR Code Absensi")

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

    st.caption("Cetak/Tempel QR ini. Karyawan scan → langsung ke form absen.")

    qr_png = build_qr_png(qr_url_effective)
    st.image(qr_png, caption="QR Absensi", use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.link_button("🔗 Tes Link Absensi", qr_url_effective, use_container_width=True)
    with c2:
        st.download_button(
            "⬇️ Download QR",
            data=qr_png,
            file_name="qr_absensi.png",
            mime="image/png",
            use_container_width=True
        )

    with st.expander("Tips penggunaan (klik untuk buka)"):
        st.write(
            "- Pastikan URL aplikasi **HTTPS** (Streamlit Cloud biasanya sudah).\n"
            "- Untuk HP jadul: jika kamera bermasalah, pakai opsi **Upload foto**.\n"
            "- Jika pakai token, QR mengandung `token=...` agar tidak sembarang orang submit."
        )
    st.stop()


# ===== PAGE: ABSEN (dibuka dari scan QR)
st.title("🧾 Form Absensi")

if ENABLE_TOKEN and TOKEN_SECRET:
    incoming_token = get_token_from_url()
    if incoming_token != TOKEN_SECRET:
        st.error("Akses tidak valid. Silakan scan QR resmi dari kantor.")
        st.stop()

dt = now_local()
ts_display = dt.strftime("%d-%m-%Y %H:%M:%S")
ts_file = dt.strftime("%Y-%m-%d_%H-%M-%S")
today_key = dt.strftime("%d-%m-%Y")

st.caption(f"🕒 Waktu server ({TZ_NAME}): **{ts_display}**")
st.info("Jika muncul pop-up izin kamera, pilih **Allow / Izinkan**. Untuk HP tertentu, gunakan **Upload foto**.")

with st.form("form_absen", clear_on_submit=False):
    st.subheader("1) Data Karyawan")

    nama = st.text_input("Nama Lengkap", placeholder="Contoh: Andi Saputra")
    no_hp = st.text_input("No HP/WA", placeholder="Contoh: 08xxxxxxxxxx atau +628xxxxxxxxxx")
    posisi = st.text_input("Posisi / Jabatan", placeholder="Contoh: Driver / Teknisi / Supervisor")

    st.divider()
    st.subheader("2) Selfie Kehadiran")

    open_cam_now = st.checkbox("Buka kamera (disarankan jika HP mendukung)", value=st.session_state.open_cam)
    st.session_state.open_cam = open_cam_now

    selfie_cam = None
    if st.session_state.open_cam:
        selfie_cam = st.camera_input("Ambil selfie")

    st.caption("Jika kamera tidak bisa dibuka, gunakan opsi upload:")
    selfie_upload = st.file_uploader("Upload foto selfie", type=["jpg", "jpeg", "png"])

    st.divider()

    submit = st.form_submit_button(
        "✅ Submit Absensi",
        disabled=st.session_state.saving or st.session_state.submitted_once,
        use_container_width=True
    )

# ===== SUBMIT LOGIC
if submit:
    if st.session_state.submitted_once:
        st.warning("Absensi sudah tersimpan. Jika ingin absen lagi, refresh halaman.")
        st.stop()

    nama_clean = sanitize_name(nama)
    hp_clean = sanitize_phone(no_hp)
    posisi_raw = str(posisi).strip()

    img_bytes, ext = get_selfie_bytes(selfie_cam, selfie_upload)

    errors = []
    if not nama_clean:
        errors.append("• Nama wajib diisi.")
    if not hp_clean or len(hp_clean.replace("+", "")) < 8:
        errors.append("• No HP/WA wajib diisi (minimal 8 digit).")
    if not posisi_raw:
        errors.append("• Posisi wajib diisi.")
    if img_bytes is None:
        errors.append("• Selfie wajib (kamera atau upload).")

    if errors:
        st.error("Mohon lengkapi dulu:\n\n" + "\n".join(errors))
        st.stop()

    st.session_state.saving = True
    try:
        with st.spinner("Menyimpan absensi..."):
            sh = connect_gsheet()
            ws = get_or_create_ws(sh)

            # anti double absen (paling penting biar rekap & download tidak meleset)
            exists, last_ts = already_checked_in_today(ws, hp_clean, today_key)
            if exists:
                st.session_state.saving = False
                st.warning(f"No HP/WA ini sudah absen hari ini (terakhir: {last_ts}). Jika itu salah, hubungi admin.")
                st.stop()

            # siapkan canonical posisi (pakai data hari ini sebagai referensi)
            try:
                hadir_today, _, _, _ = get_today_data_and_rekap()
                known_canon = [normalize_posisi(x.get("Posisi", "")) for x in hadir_today if x.get("Posisi")]
                known_canon = [k for k in known_canon if k]
            except Exception:
                known_canon = []

            posisi_norm = canonicalize_posisi(posisi_raw, known_canon=known_canon)

            dbx = connect_dropbox()
            link_selfie_raw, dbx_path = upload_selfie_to_dropbox(dbx, img_bytes, nama_clean, ts_file, ext)

            # link rapi untuk ditampilkan di Sheet
            link_cell = make_hyperlink(link_selfie_raw, "Bukti Foto")

            # append sesuai header
            ws.append_row(
                [ts_display, nama_clean, hp_clean, posisi_raw, link_cell, dbx_path, posisi_norm, link_selfie_raw],
                value_input_option="USER_ENTERED"
            )

            auto_format_absensi_sheet(ws)

            # refresh cache rekap supaya langsung update
            get_today_data_and_rekap.clear()

        st.session_state.submitted_once = True
        st.success("Absensi berhasil tersimpan. Terima kasih ✅")
        st.balloons()

        if st.button("↩️ Isi ulang (reset form)", use_container_width=True):
            st.session_state.open_cam = False
            st.session_state.saving = False
            st.session_state.submitted_once = False
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
# REKAP + DOWNLOAD (UI/UX bawah)
# =========================
st.divider()
st.subheader("📊 Rekap Kehadiran (Hari ini)")

c1, c2 = st.columns([1, 1])
with c1:
    st.caption("Rekap ini dihitung dari data Google Sheet (bukan dari form), jadi aman untuk audit.")
with c2:
    if st.button("🔄 Refresh rekap", use_container_width=True):
        get_today_data_and_rekap.clear()
        st.rerun()

try:
    hadir_today, rekap_rows, total, today_key2 = get_today_data_and_rekap()

    mc1, mc2 = st.columns([1, 1])
    with mc1:
        st.metric("Total hadir", total)
    with mc2:
        st.metric("Tanggal", today_key2)

    if total == 0:
        st.info("Belum ada absensi untuk hari ini.")
    else:
        st.table(rekap_rows)

        with st.expander("👥 Daftar hadir hari ini (klik untuk buka)"):
            # tampilkan kolom penting saja
            tampil_cols = ["Timestamp", "Nama", "No HP/WA", "Posisi"]
            st.dataframe([{k: r.get(k, "") for k in tampil_cols} for r in hadir_today], use_container_width=True)

        # DOWNLOAD CSV (biar tidak meleset)
        # 1) Rekap posisi
        rekap_csv = to_csv_bytes(rekap_rows, ["Posisi", "Jumlah"])
        st.download_button(
            "⬇️ Download Rekap Hari Ini (CSV)",
            data=rekap_csv,
            file_name=f"rekap_hadir_{today_key2}.csv",
            mime="text/csv",
            use_container_width=True
        )

        # 2) Daftar hadir detail (termasuk Selfie URL untuk bukti saat download)
        # Selfie URL diambil dari kolom "Selfie URL Raw" agar tidak salah karena formula hyperlink di sheet.
        detail_cols = ["Timestamp", "Nama", "No HP/WA", "Posisi", "Selfie URL"]
        detail_csv = to_csv_bytes(hadir_today, detail_cols)
        st.download_button(
            "⬇️ Download Daftar Hadir Hari Ini (CSV)",
            data=detail_csv,
            file_name=f"daftar_hadir_{today_key2}.csv",
            mime="text/csv",
            use_container_width=True
        )

except Exception as e:
    st.warning("Rekap kehadiran belum bisa ditampilkan (cek koneksi GSheet).")
    with st.expander("Detail error (untuk admin)"):
        st.code(str(e))
