# Altomizer

Standalone offline ALT-text management app extracted from Matcha White.
Altomizer now carries its own local processing modules under `Altomizer/`,
so the desktop app and packaged `.exe` do not depend on the web app package
at runtime.

## Desktop app

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the desktop app:

```bash
python Altomizer/run.py
```

Use the desktop app when you want the working UI to stay out of normal browser DevTools and browser menus. It runs as a native Qt window instead of a served HTML page.

What it does:

- Open a local `DOCX` or `PDF`
- Review collected ALT items in a desktop UI
- Download the ALT workbook
- Edit the workbook in Excel and import it back
- Download an updated `DOCX`
- Clear existing ALT text from a `DOCX`
- Generate numbered grid PNGs from the workbook previews

## Legacy web app

The original web prototype is still available if needed:

```bash
uvicorn Altomizer.main:app --reload
```

If you want to keep the web server up but disable the browser UI routes entirely, set:

```bash
ALTOMIZER_WEB_UI_ENABLED=0
```

When disabled, the browser routes return a minimal desktop-only notice instead of the full HTML interface.
