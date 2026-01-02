import streamlit as st
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import re
import io

import gspread
from google.oauth2.service_account import Credentials

import dropbox
from dropbox.sharing import RequestedVisibility, SharedLinkSettings
from dropbox.exceptions import ApiError, AuthError

import qrcode
from PIL import Image

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

COL_TIMESTAMP = "Timestamp"
COL_NAMA = "Nama"
COL_HP = "No HP/WA"
COL_POSISI = "Posisi"
COL_LINK_SELFIE = "Link Selfie"
COL_DBX_PATH = "Dropbox Path"

SHEET_COLUMNS = [COL_TIMESTAMP, COL_NAMA, COL_HP, COL_POSISI, COL_LINK_SELFIE, COL_DBX_PATH]


# =========================
# HELPERS
# =========================
def get_mode() -> str:
    # Kompatibel untuk versi Streamlit baru & lama
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
    # aman untuk folder/file
    text = re.sub(r"[^A-Za-z0-9 _.-]", "", text)
    return text.strip()

def sanitize_phone(text: str) -> str:
    text = str(text).strip()
    # biarkan + di awal kalau ada
    if text.startswith("+"):
        return "+" + re.sub(r"\D", "", text[1:])
    return re.sub(r"\D", "", text)

def now_jakarta():
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
    sh = gc.open(SHEET_NAME)
    return sh

def get_or_create_ws(spreadsheet):
    try:
        ws = spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=5000, cols=len(SHEET_COLUMNS))
        ws.append_row(SHEET_COLUMNS, value_input_option="USER_ENTERED")

    # pastikan header lengkap
    header = ws.row_values(1)
    if header != SHEET_COLUMNS:
        # jika kosong / beda, set ulang header
        ws.resize(cols=max(ws.col_count, len(SHEET_COLUMNS)))
        ws.update("A1", [SHEET_COLUMNS], value_input_option="USER_ENTERED")
    return ws

@st.cache_resource
def connect_dropbox():
    if "dropbox" not in st.secrets or "access_token" not in st.secrets["dropbox"]:
        raise RuntimeError("Dropbox secrets tidak ditemukan: dropbox.access_token")

    dbx = dropbox.Dropbox(st.secrets["dropbox"]["access_token"])
    dbx.users_get_current_account()
    return dbx

def upload_selfie_to_dropbox(dbx, img_bytes: bytes, nama: str, ts_file: str, ext: str) -> tuple[str, str]:
    """
    Return (shared_link_raw, dropbox_path)
    """
    clean_name = sanitize_name(nama).replace(" ", "_") or "Unknown"
    filename = f"{ts_file}_selfie{ext}"
    path = f"{DROPBOX_ROOT}/{clean_name}/{filename}"

    dbx.files_upload(img_bytes, path, mode=dropbox.files.WriteMode.add)

    settings = SharedLinkSettings(requested_visibility=RequestedVisibility.public)
    try:
        link = dbx.sharing_create_shared_link_with_settings(path, settings=settings)
        url = link.url
    except ApiError as e:
        if e.error.is_shared_link_already_exists():
            url = dbx.sharing_list_shared_links(path, direct_only=True).links[0].url
        else:
            url = "-"

    # supaya bisa langsung preview image
    url_raw = url.replace("?dl=0", "?raw=1") if url and url != "-" else "-"
    return url_raw, path


# =========================
# UI
# =========================
mode = get_mode()

# ===== PAGE: QR / ADMIN
if mode != "absen":
    st.title("✅ QR Code Absensi (Statis)")

    if not QR_URL:
        st.warning("QR URL belum diset. Isi `app.qr_url` di secrets.")
        st.code("Contoh: https://YOUR-APP.streamlit.app/?mode=absen", language="text")
        st.stop()

    # opsional: token
    if ENABLE_TOKEN and TOKEN_SECRET:
        # kalau qr_url belum ada token, tambahkan otomatis untuk preview
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

    st.link_button("🔗 Buka Link Absensi (Tes)", qr_url_effective, use_container_width=True)
    st.download_button(
        "⬇️ Download QR PNG",
        data=qr_png,
        file_name="qr_absensi.png",
        mime="image/png",
        use_container_width=True
    )

    st.divider()
    st.info(
        "Tips:\n"
        "- Setelah deploy, pastikan `app.qr_url` sudah pakai URL Streamlit Cloud Anda.\n"
        "- Jika pakai token, QR akan mengandung `token=...` agar tidak sembarang orang submit."
    )
    st.stop()


# ===== PAGE: ABSEN (dibuka dari scan QR)
st.title("🧾 Form Absensi (Scan QR)")

if ENABLE_TOKEN and TOKEN_SECRET:
    incoming_token = get_token_from_url()
    if incoming_token != TOKEN_SECRET:
        st.error("Akses tidak valid. Silakan scan QR resmi dari kantor.")
        st.stop()

# Tampilkan timestamp server
dt = now_jakarta()
ts_display = dt.strftime("%d-%m-%Y %H:%M:%S")
ts_file = dt.strftime("%Y-%m-%d_%H-%M-%S")

st.caption(f"🕒 Waktu server (Asia/Jakarta): **{ts_display}**")

nama = st.text_input("Nama Lengkap", placeholder="Contoh: Andi Saputra")
no_hp = st.text_input("No HP/WA", placeholder="Contoh: 08xxxxxxxxxx atau +628xxxxxxxxxx")
posisi = st.text_input("Posisi / Jabatan", placeholder="Contoh: Sales / Admin / Marketing")

st.divider()
st.subheader("🤳 Selfie Kehadiran")
selfie = st.camera_input("Ambil selfie (kamera HP akan terbuka)")

submit = st.button("✅ Submit Absensi", type="primary", use_container_width=True)

if submit:
    # Validasi
    nama_clean = sanitize_name(nama)
    hp_clean = sanitize_phone(no_hp)
    posisi_clean = str(posisi).strip()

    if not nama_clean:
        st.error("Nama wajib diisi.")
        st.stop()
    if not hp_clean or len(hp_clean.replace("+", "")) < 8:
        st.error("No HP/WA wajib diisi (minimal 8 digit).")
        st.stop()
    if not posisi_clean:
        st.error("Posisi wajib diisi.")
        st.stop()
    if selfie is None:
        st.error("Selfie wajib diambil.")
        st.stop()

    # Tentukan ekstensi file dari mime type
    ext = ".jpg"
    if getattr(selfie, "type", "") == "image/png":
        ext = ".png"

    try:
        with st.spinner("Menyimpan absensi..."):
            # Connect
            sh = connect_gsheet()
            ws = get_or_create_ws(sh)

            dbx = connect_dropbox()

            # Upload selfie
            img_bytes = selfie.getvalue()
            link_selfie, dbx_path = upload_selfie_to_dropbox(dbx, img_bytes, nama_clean, ts_file, ext)

            # Append row to sheet
            ws.append_row(
                [ts_display, nama_clean, hp_clean, posisi_clean, link_selfie, dbx_path],
                value_input_option="USER_ENTERED"
            )

        st.success("Absensi berhasil tersimpan. Terima kasih ✅")
        if link_selfie and link_selfie != "-":
            st.link_button("🔎 Lihat Selfie (Dropbox)", link_selfie, use_container_width=True)

        st.balloons()

    except AuthError:
        st.error("Dropbox token tidak valid. Hubungi admin.")
    except Exception as e:
        st.error(f"Gagal menyimpan absensi: {e}")
