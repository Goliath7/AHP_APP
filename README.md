# Document Generator (Web)

Web app for generating 3 types of documents:
- Outgoing letter (`Исходящее письмо`)
- Work permit (`Наряд`)
- Work plan (`План работ`)

The web UI is built with Streamlit and can be deployed for free on Streamlit Community Cloud.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Project files

- `streamlit_app.py` - main web interface
- `letter_doc_builder.py` - outgoing letter generator (no tkinter)
- `nariad_gui_final.py` - work permit generator (reused by web app)
- `plan_rabot_GUI.py` - work plan generator (reused by web app)
- `shared_data.py` - shared stations/leaders/output path utilities

## Free deployment (Streamlit Community Cloud)

1. Push this folder to a GitHub repository.
2. Open [share.streamlit.io](https://share.streamlit.io/).
3. Click `New app`.
4. Select repository and set:
   - Branch: `main`
   - Main file path: `streamlit_app.py`
5. Click `Deploy`.

## Quick GitHub push commands

Run in this folder:

```bash
git init
git add .
git commit -m "Initial web app"
git branch -M main
git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPO>.git
git push -u origin main
```
