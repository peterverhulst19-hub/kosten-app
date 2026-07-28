import io
from datetime import date

import openpyxl
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

TRANSACTIONS_SHEET = "Transacties"
COLUMNS = ["Datum", "Jaar", "Maand", "Categorie", "Subcategorie", "Bedrag", "Omschrijving"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

DUTCH_MONTHS = [
    "Januari", "Februari", "Maart", "April", "Mei", "Juni",
    "Juli", "Augustus", "September", "Oktober", "November", "December",
]

# Best-effort mapping detected from the example workbook. Adjust freely if it
# doesn't match how you actually want expenses grouped.
CATEGORY_MAP = {
    "Voeding": ["Voeding winkel", "Uit eten"],
    "Vaste kosten": [
        "Leningen en verzekeringen", "Sparen", "Internet en tv",
        "Nutsvoorzieningen", "Apotheek en dokter", "Kinderen school",
    ],
    "Niet vaste kosten": [
        "Kleding en accessoires", "Woning en tuin", "Huisdieren", "Hobbies",
        "Kinderen buitenschoolse activiteiten", "Kinderen speelgoed en feesten",
        "Vakantie", "Feesten en gelegenheden",
    ],
    "Vervoer": ["Vervoer"],
}

# --- Mapping onto the per-year "Vaste kosten {jaar}" sheet layout ---
# Derived from inspecting the example workbook's month-block grid (12 blocks of
# 3 columns each, starting at column A). Adjust here if the real sheet differs.
YEAR_SHEET_TEMPLATE = "Vaste kosten {jaar}"
GRAND_TOTAL_ROW = 4
TOP_LEVEL_ROW = {"Voeding": 5, "Vaste kosten": 6, "Niet vaste kosten": 7, "Vervoer": 8}

# subcategorie -> (eerste data-rij, laatste data-rij) binnen het blok onder het label
# "Voeding winkel" en "Uit eten" hebben in de template geen eigen invoerrijen; voor
# hen wordt enkel de hoofdcategorie-totaal bijgewerkt.
SUBCATEGORY_BLOCKS = {
    "Leningen en verzekeringen": (12, 17),
    "Sparen": (19, 21),
    "Internet en tv": (23, 26),
    "Nutsvoorzieningen": (28, 30),
    "Apotheek en dokter": (32, 39),
    "Kinderen school": (50, 57),
    "Vervoer": (59, 65),
    "Woning en tuin": (69, 78),
    "Kleding en accessoires": (80, 82),
    "Huisdieren": (84, 86),
    "Hobbies": (88, 92),
    "Kinderen buitenschoolse activiteiten": (94, 98),
    "Kinderen speelgoed en feesten": (100, 103),
    "Vakantie": (105, 120),
    "Feesten en gelegenheden": (122, 131),
}


def month_base_column(maand_index: int) -> int:
    return 1 + 3 * maand_index


def update_year_sheet(wb, jaar: int, maand_index: int, hoofdcategorie: str,
                       subcategorie: str, bedrag: float, omschrijving: str) -> list:
    warnings = []
    sheet_name = YEAR_SHEET_TEMPLATE.format(jaar=jaar)
    if sheet_name not in wb.sheetnames:
        warnings.append(f"Sheet '{sheet_name}' niet gevonden, enkel toegevoegd aan {TRANSACTIONS_SHEET}.")
        return warnings

    ws = wb[sheet_name]
    base_col = month_base_column(maand_index)
    amount_col = base_col + 1

    for row in (GRAND_TOTAL_ROW, TOP_LEVEL_ROW[hoofdcategorie]):
        cell = ws.cell(row=row, column=amount_col)
        cell.value = (cell.value or 0) + bedrag

    block = SUBCATEGORY_BLOCKS.get(subcategorie)
    if block is None:
        return warnings

    first_row, last_row = block
    for r in range(first_row, last_row + 1):
        if ws.cell(row=r, column=amount_col).value in (None, ""):
            ws.cell(row=r, column=amount_col).value = bedrag
            ws.cell(row=r, column=amount_col + 1).value = omschrijving
            break
    else:
        warnings.append(
            f"Geen vrije rij meer in het blok '{subcategorie}' van '{sheet_name}' "
            f"(totalen zijn wel bijgewerkt)."
        )
    return warnings


# --- Google Drive opslag ---
# Het bestand blijft altijd op Drive; de app haalt het op, wijzigt het in het
# geheugen en zet het meteen terug weg. Zo is er geen lokale schijf nodig en
# werkt dit ook vanaf de cloud.

@st.cache_resource
def get_drive_service():
    info = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive.file"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


GOOGLE_SHEETS_MIME = "application/vnd.google-apps.spreadsheet"


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


def load_transactions(file_id: str) -> pd.DataFrame:
    try:
        data = download_workbook_bytes(file_id)
    except Exception as exc:
        st.error(f"Kon bestand niet ophalen van Google Drive: {exc}")
        return pd.DataFrame(columns=COLUMNS)

    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
    sheet_names = wb.sheetnames
    wb.close()
    if TRANSACTIONS_SHEET not in sheet_names:
        return pd.DataFrame(columns=COLUMNS)

    df = pd.read_excel(io.BytesIO(data), sheet_name=TRANSACTIONS_SHEET)
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[COLUMNS]


def save_expense(file_id: str, row: dict, maand_index: int) -> list:
    data = download_workbook_bytes(file_id)
    wb = openpyxl.load_workbook(io.BytesIO(data))

    if TRANSACTIONS_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(TRANSACTIONS_SHEET)
        ws.append(COLUMNS)
    else:
        ws = wb[TRANSACTIONS_SHEET]
    ws.append([row[col] for col in COLUMNS])

    warnings = update_year_sheet(
        wb, row["Jaar"], maand_index, row["Categorie"], row["Subcategorie"],
        row["Bedrag"], row["Omschrijving"],
    )

    out = io.BytesIO()
    wb.save(out)
    upload_workbook_bytes(file_id, out.getvalue())
    return warnings


st.set_page_config(page_title="Kosten invoer", page_icon="\U0001F4B0")

# --- Configuratie-check: geeft een duidelijke melding i.p.v. een kale
# "Internal server error" zolang niet alle secrets zijn ingevuld. ---
REQUIRED_SECRETS = ["auth", "gcp_service_account", "drive_file_id"]
missing_secrets = [key for key in REQUIRED_SECRETS if key not in st.secrets]
if missing_secrets:
    st.title("Kosten invoer")
    st.error(
        "De app is nog niet volledig geconfigureerd. Ontbrekende secrets: "
        + ", ".join(missing_secrets)
        + ". Voeg deze toe via Settings → Secrets op Streamlit Cloud."
    )
    st.stop()

# --- Login (enkel toegankelijk voor jezelf) ---
if not st.user.is_logged_in:
    st.title("Kosten invoer")
    st.write("Log in met Google om je kosten te beheren.")
    st.button("Inloggen met Google", on_click=st.login)
    st.stop()

allowed_emails = st.secrets.get("allowed_emails", [])
if allowed_emails and st.user.email not in allowed_emails:
    st.error(f"Geen toegang voor {st.user.email}.")
    st.button("Uitloggen", on_click=st.logout)
    st.stop()

FILE_ID = st.secrets["drive_file_id"]

st.title("Kosten invoer")
with st.sidebar:
    st.caption(f"Ingelogd als {st.user.email}")
    st.button("Uitloggen", on_click=st.logout)

hoofdcategorie = st.selectbox("Hoofdcategorie", list(CATEGORY_MAP.keys()))

with st.form("nieuwe_kost", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        datum = st.date_input("Datum", value=date.today())
        subcategorieen = CATEGORY_MAP[hoofdcategorie]
        subcategorie = st.selectbox("Subcategorie", subcategorieen) if subcategorieen else ""
    with col2:
        bedrag = st.number_input("Bedrag (EUR)", min_value=0.0, step=0.5, format="%.2f")
        omschrijving = st.text_input("Omschrijving (optioneel)")
    submitted = st.form_submit_button("Toevoegen")

    if submitted:
        if bedrag <= 0:
            st.error("Vul een bedrag groter dan 0 in.")
        else:
            row = {
                "Datum": datum,
                "Jaar": datum.year,
                "Maand": DUTCH_MONTHS[datum.month - 1],
                "Categorie": hoofdcategorie,
                "Subcategorie": subcategorie,
                "Bedrag": bedrag,
                "Omschrijving": omschrijving,
            }
            try:
                warnings = save_expense(FILE_ID, row, datum.month - 1)
            except Exception as exc:
                st.error(f"Kon niet opslaan naar Google Drive: {exc}")
            else:
                st.success(f"Toegevoegd: {hoofdcategorie} / {subcategorie or '-'} - EUR {bedrag:.2f}")
                for warning in warnings:
                    st.warning(warning)

st.divider()
st.subheader("Overzicht")

transacties = load_transactions(FILE_ID)

if transacties.empty:
    st.info("Nog geen kosten ingevoerd.")
else:
    jaren = sorted(transacties["Jaar"].dropna().unique(), reverse=True)
    jaar = st.selectbox("Jaar", jaren)
    categorieen = st.multiselect(
        "Categorie filter", sorted(transacties["Categorie"].dropna().unique())
    )

    gefilterd = transacties[transacties["Jaar"] == jaar]
    if categorieen:
        gefilterd = gefilterd[gefilterd["Categorie"].isin(categorieen)]

    st.dataframe(
        gefilterd.sort_values("Datum", ascending=False),
        use_container_width=True,
        hide_index=True,
    )

    if not gefilterd.empty:
        st.subheader(f"Samenvatting {jaar}")
        pivot = pd.pivot_table(
            gefilterd, index="Maand", columns="Categorie", values="Bedrag",
            aggfunc="sum", fill_value=0,
        ).reindex(DUTCH_MONTHS).dropna(how="all")
        st.dataframe(pivot, use_container_width=True)

        per_categorie = gefilterd.groupby("Categorie")["Bedrag"].sum()
        st.bar_chart(per_categorie)
