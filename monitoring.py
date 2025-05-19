import streamlit as st
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Ganti dengan ID spreadsheet dan nama sheet Anda
SPREADSHEET_ID = '1yVTIWSOBz22XaTkF6iYD90HuYXJvdzqiGaEcW-K6l6g'
DATA_SPK = "data_spk"
HM_HARIAN = "HM_Harian"
DATA_OLI = "data_oli"
PENGGUNAAN_FK = "penggunaan_forklift"

# Scope untuk Google Sheets API
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Ambil kredensial dari Streamlit Secrets
creds_info = st.secrets["google_service_account"]

# Autentikasi dan bangun service
@st.cache_resource
def get_service():
    try:
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        return build('sheets', 'v4', credentials=creds)
    except Exception as e:
        st.error(f"Gagal terhubung ke Google Sheets: {e}")
        st.stop()

service = get_service()

# --- Membaca Data dari Sheet (dengan cache) ---
@st.cache_data(ttl=28800, show_spinner=False)  # cache selama 1 jam, bisa diubah sesuai kebutuhan
def read_sheet(SHEET_NAME):
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}"
        ).execute()
        values = result.get('values', [])
        df = pd.DataFrame(values[1:], columns=values[0]) if values else pd.DataFrame()
        return df
    except Exception as e:
        st.error(f"Tidak dapat membaca data dari sheet '{SHEET_NAME}': {e}")
        return pd.DataFrame()  # Kembalikan dataframe kosong sebagai fallback

# --- Menulis Data ke Sheet ---
def write_to_sheet(SHEET_NAME, dataframe):
    try:
        service.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID,
            range=SHEET_NAME
        ).execute()
        
        values = [dataframe.columns.tolist()] + dataframe.values.tolist()
        body = {'values': values}
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=SHEET_NAME,
            valueInputOption="RAW",
            body=body
        ).execute()

        read_sheet.clear()
    except Exception as e:
        st.error(f"Gagal menyimpan data ke sheet '{SHEET_NAME}': {e}")


df_hmharian = read_sheet(HM_HARIAN)

# Baca file HM_Harian.xlsx
all_columns = df_hmharian.columns.tolist()

forklifts = sorted({
    col.split(" - ")[0]
    for col in all_columns
    if " - " in col
})

# Buat dict estimasi
estimasi_dict = {}

for fk in forklifts:
    # Siapkan nama-nama kolom shift
    shift_cols = [f"{fk} - Shift {i}" for i in (1, 2, 3)]
    last_hm = None

    # Loop mundur dari baris paling bawah (tanggal paling akhir)  
    for idx in reversed(df_hmharian.index):
        row = df_hmharian.loc[idx]

        # Cek Shift 3 → Shift 2 → Shift 1
        for shift in (3, 2, 1):
            col = f"{fk} - Shift {shift}"
            if col not in df_hmharian.columns:
                continue

            # Coba konversi ke numerik
            val = row[col]
            num = pd.to_numeric(val, errors="coerce")

            if pd.notnull(num):
                last_hm = num
                break  # keluar loop shift

        if last_hm is not None:
            break  # keluar loop tanggal

    # Jika ketemu nilai, tambahkan 21
    estimasi_dict[fk] = (last_hm + 21) if last_hm is not None else None

# --- Langkah 2: Baca data_oli dan tambahkan kolom Estimasi HM Hari Ini ---
df_oli = read_sheet(DATA_OLI)
df_oli["Estimasi HM Hari Ini"] = df_oli["No. FK"].map(estimasi_dict)

# Konversi ke numerik dan isi NaN dengan 0
for col in [
    "HM Terakhir Ganti Oli mesin",
    "HM Terakhir Ganti Oli Hidrolik",
    "HM Terakhir Saat Ganti Oli Transmisi",
    "HM Terakhir Saat Ganti Oli Gardan",
    "Estimasi HM Hari Ini"
]:
    df_oli[col] = pd.to_numeric(df_oli[col], errors="coerce").fillna(0)

# --- Langkah 3: Menghitung Sisa HM Ganti Oli ---
df_oli["Sisa HM Ganti Oli mesin"]      = df_oli["HM Terakhir Ganti Oli mesin"]      + 250  - df_oli["Estimasi HM Hari Ini"]
df_oli["Sisa HM Ganti Oli Hidrolik"]   = df_oli["HM Terakhir Ganti Oli Hidrolik"]   + 3000 - df_oli["Estimasi HM Hari Ini"]
df_oli["Sisa HM Ganti Oli Transmisi"]  = df_oli["HM Terakhir Saat Ganti Oli Transmisi"] + 2500 - df_oli["Estimasi HM Hari Ini"]
df_oli["Sisa HM Ganti Oli Gardan"]     = df_oli["HM Terakhir Saat Ganti Oli Gardan"]    + 2500 - df_oli["Estimasi HM Hari Ini"]

# --- Langkah 4: Tampilkan Tabel Non-editable ---
non_editable_columns = [
    "No. FK", "Status", "Estimasi HM Hari Ini",
    "Sisa HM Ganti Oli mesin", "Sisa HM Ganti Oli Hidrolik",
    "Sisa HM Ganti Oli Transmisi", "Sisa HM Ganti Oli Gardan"
]

def highlight_red(val):
    try:
        v = float(val)
        if v <= 63:
            return "background-color: red"
        elif v <= 147:
            return "background-color: yellow"
    except:
        pass
    return ""

styled_df = (
    df_oli[non_editable_columns]
    .style
    .format({
        "Estimasi HM Hari Ini": "{:.0f}",
        "Sisa HM Ganti Oli mesin": "{:.0f}",
        "Sisa HM Ganti Oli Hidrolik": "{:.0f}",
        "Sisa HM Ganti Oli Transmisi": "{:.0f}",
        "Sisa HM Ganti Oli Gardan": "{:.0f}",
    })
    .applymap(highlight_red, subset=[
        "Sisa HM Ganti Oli mesin",
        "Sisa HM Ganti Oli Hidrolik",
        "Sisa HM Ganti Oli Transmisi",
        "Sisa HM Ganti Oli Gardan",
    ])
)

st.subheader("Tabel Estimasi HM dan Sisa HM Ganti Oli")
st.write(styled_df)
