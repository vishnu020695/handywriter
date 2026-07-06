# HandyWriter — Free Handwriting-to-Word/Excel Converter + PDF Editor

100% free and open-source. No subscriptions, no accounts, no internet required
after setup. Your files never leave your device.

## What it does
1. **Image → Word/Excel**: Upload a photo of handwriting or a scanned page.
   It reads the text automatically, lets you correct any mistakes, then saves
   it as a `.docx` or `.xlsx` file.
2. **PDF Editor**: Delete pages, rotate pages, add a watermark, or convert a
   PDF into an editable Word document (keeps the original layout/alignment).

⚠️ **Honesty note:** No OCR tool — free or paid — reads messy handwriting
with 100% accuracy. This app always shows you the extracted text in an
editable box before saving, so you can fix anything it got wrong.

---

## Install on Windows / Mac / Linux (Desktop or Laptop)

**Step 1 — Install Python** (only needed once):
- Go to https://www.python.org/downloads/ and install Python 3.10 or newer.
- On Windows, tick "Add Python to PATH" during install.

**Step 2 — Install Tesseract OCR** (the free text-reading engine):
- Windows: download the installer from
  https://github.com/UB-Mannheim/tesseract/wiki
- Mac: open Terminal → `brew install tesseract`
- Linux: `sudo apt install tesseract-ocr`

**Step 3 — Download this app's files** (`app.py` and `requirements.txt`)
into one folder.

**Step 4 — Open a terminal / command prompt in that folder and run:**
```
pip install -r requirements.txt
streamlit run app.py
```

Your browser will open automatically to `http://localhost:8501` — that's
the app. To use it again later, just repeat Step 4 (no need to reinstall).

---

## Use it on your Phone (Android or iPhone)

The easiest way — no app-store install needed:

1. Set it up on your laptop/desktop as above, and run `streamlit run app.py`.
2. Make sure your phone is on the **same WiFi network** as your computer.
3. In the terminal where Streamlit is running, note the "Network URL" it
   prints (looks like `http://192.168.1.23:8501`).
4. Open that address in your phone's browser (Chrome/Safari). The app now
   works on your phone too — you can even use your phone's camera to snap
   the handwritten photo directly into the upload box.

### Advanced: running Python directly on Android (no laptop needed)
Install the free **Termux** app from F-Droid (https://f-droid.org/packages/com.termux/),
then inside Termux run:
```
pkg install python tesseract
pip install -r requirements.txt
streamlit run app.py
```
This runs the whole thing on your phone alone.

---

## Files in this project
- `app.py` — the application
- `requirements.txt` — the list of free Python libraries it needs

## Tools used (all free & open-source)
- Tesseract OCR — reads text from images
- python-docx — creates Word documents
- openpyxl — creates Excel spreadsheets
- PyMuPDF (fitz) — edits/rotates/watermarks PDFs
- pdf2docx — converts PDF layout to editable Word
- Streamlit — the easy-to-use screen/interface
