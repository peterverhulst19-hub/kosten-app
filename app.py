import io
from datetime import date

import folium
import openpyxl
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from streamlit_folium import st_folium
from streamlit_js_eval import get_geolocation

LOCATIONS_SHEET = "Locaties"
COLUMNS = ["Datum", "Naam", "Notities", "Latitude", "Longitude"]
PHOTO_COLUMN = "F"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GOOGLE_SHEETS_MIME = "application/vnd.google-apps.spreadsheet"

# Antwerpen, gebruikt als startpunt zolang er nog geen locatie gekozen is.
DEFAULT_LAT, DEFAULT_LON = 51.2194, 4.4025


def compress_photo(data: bytes, max_size: int = 1024, quality: int = 75) -> bytes:
    img = PILImage.open(io.BytesIO(data)).convert("RGB")
    img.thumbnail((max_size, max_size))
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=quality)
    return out.getvalue()


def image_bytes(xl_image) -> bytes:
    try:
        return xl_image._data()
    except Exception:
        buf = io.BytesIO()
        xl_image.ref.save(buf, format="PNG")
        return buf.getvalue()


# --- Google Drive opslag ---
# Het bestand blijft altijd op Drive; de app haalt het op, wijzigt het in het
# geheugen en zet het meteen terug weg. Zo is er geen lokale schijf nodig en
# werkt dit ook vanaf de cloud.

@st.cache_resource
def get_drive_service():
    info = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def download_workbook_bytes(file_id: str) -> bytes:
    service = get_drive_service()
    meta = service.files().get(fileId=file_id, fields="mimeType").execute()
    if meta["mimeType"] == GOOGLE_SHEETS_MIME:
        # Native Google Sheet (aangemaakt via "Nieuw -> Google Spreadsheets"):
        # moet geëxporteerd worden als xlsx, kan niet rechtstreeks gedownload worden.
        request = service.files().export_media(fileId=file_id, mimeType=XLSX_MIME)
    else:
        # Een echt geüpload .xlsx-bestand.
        request = service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def upload_workbook_bytes(file_id: str, data: bytes) -> None:
    service = get_drive_service()
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=XLSX_MIME, resumable=False)
    service.files().update(fileId=file_id, media_body=media).execute()


def save_location(file_id: str, row: dict, photo_bytes: bytes | None) -> None:
    data = download_workbook_bytes(file_id)
    wb = openpyxl.load_workbook(io.BytesIO(data))  # niet read_only: nodig om afbeeldingen toe te voegen

    if LOCATIONS_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(LOCATIONS_SHEET)
        ws.append(COLUMNS)
    else:
        ws = wb[LOCATIONS_SHEET]
    ws.append([row[col] for col in COLUMNS])
    row_num = ws.max_row

    if photo_bytes:
        pil_img = PILImage.open(io.BytesIO(photo_bytes))
        target_width = 160
        target_height = int(pil_img.height * (target_width / pil_img.width))
        xl_image = XLImage(io.BytesIO(photo_bytes))
        xl_image.width, xl_image.height = target_width, target_height
        ws.add_image(xl_image, f"{PHOTO_COLUMN}{row_num}")
        ws.row_dimensions[row_num].height = target_height * 0.75  # px -> punten

    out = io.BytesIO()
    wb.save(out)
    upload_workbook_bytes(file_id, out.getvalue())


def load_locations(file_id: str):
    try:
        data = download_workbook_bytes(file_id)
    except Exception as exc:
        st.error(f"Kon bestand niet ophalen van Google Drive: {exc}")
        return pd.DataFrame(columns=COLUMNS), {}

    wb = openpyxl.load_workbook(io.BytesIO(data))  # niet read_only: nodig om afbeeldingen te lezen
    if LOCATIONS_SHEET not in wb.sheetnames:
        return pd.DataFrame(columns=COLUMNS), {}

    ws = wb[LOCATIONS_SHEET]
    df = pd.read_excel(io.BytesIO(data), sheet_name=LOCATIONS_SHEET)
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df[COLUMNS]

    row_photos = {}
    for xl_image in ws._images:
        sheet_row = xl_image.anchor._from.row + 1  # 0-indexed -> 1-indexed
        df_index = sheet_row - 2  # rij 1 = header, rij 2 = df-index 0
        if df_index >= 0:
            row_photos[df_index] = image_bytes(xl_image)
    return df, row_photos


st.set_page_config(page_title="Leuke locaties", page_icon="\U0001F4CD")

# --- Configuratie-check: geeft een duidelijke melding i.p.v. een kale
# "Internal server error" zolang niet alle secrets zijn ingevuld. ---
REQUIRED_SECRETS = ["auth", "gcp_service_account", "drive_file_id"]
missing_secrets = [key for key in REQUIRED_SECRETS if key not in st.secrets]
if missing_secrets:
    st.title("Leuke locaties")
    st.error(
        "De app is nog niet volledig geconfigureerd. Ontbrekende secrets: "
        + ", ".join(missing_secrets)
        + ". Voeg deze toe via Settings → Secrets op Streamlit Cloud."
    )
    st.stop()

# --- Login (enkel toegankelijk voor jezelf) ---
if not st.user.is_logged_in:
    st.title("Leuke locaties")
    st.write("Log in met Google om je locaties te beheren.")
    st.button("Inloggen met Google", on_click=st.login)
    st.stop()

allowed_emails = st.secrets.get("allowed_emails", [])
if allowed_emails and st.user.email not in allowed_emails:
    st.error(f"Geen toegang voor {st.user.email}.")
    st.button("Uitloggen", on_click=st.logout)
    st.stop()

FILE_ID = st.secrets["drive_file_id"]

st.title("Leuke locaties")
with st.sidebar:
    st.caption(f"Ingelogd als {st.user.email}")
    st.button("Uitloggen", on_click=st.logout)

st.subheader("Nieuwe locatie toevoegen")

if "lat" not in st.session_state:
    st.session_state.lat = DEFAULT_LAT
    st.session_state.lon = DEFAULT_LON

geoloc = get_geolocation()
if st.button("📍 Gebruik mijn huidige locatie"):
    if geoloc and "coords" in geoloc:
        st.session_state.lat = geoloc["coords"]["latitude"]
        st.session_state.lon = geoloc["coords"]["longitude"]
    elif geoloc and "error" in geoloc:
        st.warning(
            f"Kon je locatie niet ophalen: {geoloc['error'].get('message', 'onbekende fout')}. "
            "Geef de browser toestemming voor locatietoegang en probeer opnieuw."
        )
    else:
        st.warning("Nog geen locatie beschikbaar, probeer het nog eens.")

picker_map = folium.Map(location=[st.session_state.lat, st.session_state.lon], zoom_start=13)
folium.Marker([st.session_state.lat, st.session_state.lon]).add_to(picker_map)
map_click = st_folium(picker_map, height=350, use_container_width=True, key="location_picker")
if map_click and map_click.get("last_clicked"):
    st.session_state.lat = map_click["last_clicked"]["lat"]
    st.session_state.lon = map_click["last_clicked"]["lng"]

st.caption(f"Geselecteerde locatie: {st.session_state.lat:.5f}, {st.session_state.lon:.5f}")

with st.form("nieuwe_locatie", clear_on_submit=True):
    naam = st.text_input("Naam")
    notities = st.text_area("Notities (optioneel)")
    datum = st.date_input("Datum", value=date.today())

    foto_modus = st.radio("Foto", ["Geen", "Nemen", "Kiezen"], horizontal=True)
    foto_ruw = None
    if foto_modus == "Nemen":
        camera_bestand = st.camera_input("Neem een foto")
        if camera_bestand:
            foto_ruw = camera_bestand.getvalue()
    elif foto_modus == "Kiezen":
        upload_bestand = st.file_uploader("Kies een foto", type=["jpg", "jpeg", "png"])
        if upload_bestand:
            foto_ruw = upload_bestand.getvalue()

    submitted = st.form_submit_button("Toevoegen")

    if submitted:
        if not naam:
            st.error("Vul een naam in.")
        else:
            row = {
                "Datum": datum,
                "Naam": naam,
                "Notities": notities,
                "Latitude": st.session_state.lat,
                "Longitude": st.session_state.lon,
            }
            foto_bytes = compress_photo(foto_ruw) if foto_ruw else None
            try:
                save_location(FILE_ID, row, foto_bytes)
            except Exception as exc:
                st.error(f"Kon niet opslaan naar Google Drive: {exc}")
            else:
                st.success(f"'{naam}' toegevoegd!")

st.divider()
st.subheader("Overzicht")

locaties, row_photos = load_locations(FILE_ID)

if locaties.empty:
    st.info("Nog geen locaties toegevoegd.")
else:
    zoek = st.text_input("Zoeken", placeholder="Naam of notities...")
    gefilterd = locaties
    if zoek:
        mask = (
            locaties["Naam"].str.contains(zoek, case=False, na=False)
            | locaties["Notities"].str.contains(zoek, case=False, na=False)
        )
        gefilterd = locaties[mask]

    if not gefilterd.empty:
        overview_map = folium.Map(
            location=[gefilterd["Latitude"].mean(), gefilterd["Longitude"].mean()],
            zoom_start=7,
        )
        for _, r in gefilterd.iterrows():
            folium.Marker(
                [r["Latitude"], r["Longitude"]],
                popup=f"{r['Naam']} ({r['Datum']})",
            ).add_to(overview_map)
        st_folium(overview_map, height=350, use_container_width=True, key="overview_map")

    for idx, r in gefilterd.sort_values("Datum", ascending=False).iterrows():
        with st.container(border=True):
            cols = st.columns([1, 2])
            with cols[0]:
                if idx in row_photos:
                    st.image(row_photos[idx])
            with cols[1]:
                st.markdown(f"**{r['Naam']}** — {r['Datum']}")
                if r.get("Notities"):
                    st.write(r["Notities"])
                maps_url = f"https://www.google.com/maps?q={r['Latitude']},{r['Longitude']}"
                st.markdown(f"[📍 Open in Google Maps]({maps_url})")
