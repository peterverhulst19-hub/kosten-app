# Leuke locaties

A private Streamlit app for tracking your favorite places — restaurants, sights,
museums, places to stay, walks, ... — on a map, with notes and photos.

## Features

- **Add a location** by using your current GPS position or by clicking a spot on the map.
- **Categorize** each entry (🍽️ Eten en drinken, 🏰 Bezichtigingen, 🖼️ Musea, 🛏️ Slaaplocaties,
  🥾 Wandelingen, 📍 Overig), with a per-category map and tab.
- **Attach a photo**, taken with the camera or uploaded, shown on the card and on the map popup.
- **Search** locations by name or notes.
- **Edit or delete** any entry afterwards.

## How data is stored

- Locations live in a "Locaties" sheet inside an Excel workbook on Google Drive
  (`drive_file_id` secret). The app downloads it, applies the change, and re-uploads it —
  no local database.
- Photos are uploaded to a folder named "Leuke locaties - fotos" in the *logged-in user's own*
  Google Drive (via their OAuth token), since the Google service account has no storage quota
  of its own.

## Access

- Login is via Google (`st.login`); only Google accounts listed in the `allowed_emails` secret
  can use the app (if that list is set).
- Sessions expire automatically after 30 minutes of inactivity.

## Required secrets (`.streamlit/secrets.toml`)

```toml
drive_file_id = "..."          # Google Drive file ID of the Excel workbook
allowed_emails = ["you@example.com"]   # optional allow-list

[auth]
redirect_uri = "..."
cookie_secret = "..."
client_id = "..."
client_secret = "..."
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
expose_tokens = true            # required so photo uploads can use the user's Drive token
client_kwargs = { scope = "openid email profile https://www.googleapis.com/auth/drive.file" }

[gcp_service_account]
# standard Google service-account JSON key fields
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "..."
client_email = "..."
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
```

The service account only needs access to the workbook itself (share the file with its
`client_email`); it's used to read/write the sheet, not the photos.

## Local development

```
pip install -r requirements.txt
streamlit run app.py
```

## Deployment

Deployed on Streamlit Community Cloud with "Main file path" set to `leuke-locaties/app.py`.
