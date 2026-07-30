# streamlit

This repo hosts multiple independent Streamlit apps, one per subfolder. Each app has its own `requirements.txt`, `runtime.txt`, and `.streamlit/` config, so they can be deployed separately on Streamlit Community Cloud by pointing the app's "Main file path" at the right subfolder (e.g. `locaties/app.py`).

## Apps

- [`locaties/`](locaties/) — Uitgaven doorheen het jaar (expense/location tracking app).
- [`plants/`](plants/) — Compare plant species prices across webshops and physical plant centers.

## Local development

Each app is run from its own folder, e.g.:

```
pip install -r locaties/requirements.txt
streamlit run locaties/app.py
```
