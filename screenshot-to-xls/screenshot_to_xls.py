"""
screenshot_to_xls.py
--------------------
Monitors clipboard (PrtScn) and Screenshots folder (Win+PrtScn).
Extracts text via OCR and saves everything to an Excel file.

Usage:
    python screenshot_to_xls.py
    python screenshot_to_xls.py --output my_data.xlsx
    python screenshot_to_xls.py --folder "C:/Users/YourName/Pictures/Screenshots"
"""

import os
import sys
import time
import hashlib
import argparse
import threading
from datetime import datetime

from PIL import ImageGrab, Image
import pytesseract
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl import load_workbook
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ─── Config ──────────────────────────────────────────────────────────────────

DEFAULT_XLS      = r"D:\screenshots_data.xlsx"
DEFAULT_SHOTS_DIR = os.path.expanduser("~/Pictures/Screenshots")
TESSERACT_PATH   = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Set Tesseract path (Windows)
if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# ─── Excel Setup ─────────────────────────────────────────────────────────────

def init_excel(filepath):
    """Create Excel file with styled headers if it doesn't exist."""
    if not os.path.exists(filepath):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Screenshots"

        headers = ["#", "Timestamp", "Source", "Filename", "Extracted Text"]
        ws.append(headers)

        # Style header row
        header_fill = PatternFill("solid", fgColor="1F4E79")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill   = header_fill
            cell.font   = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Column widths
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 20
        ws.column_dimensions["D"].width = 30
        ws.column_dimensions["E"].width = 80

        ws.row_dimensions[1].height = 20
        wb.save(filepath)
        print(f"📊 Created Excel file: {filepath}")
    else:
        ws = load_workbook(filepath).active
        existing_rows = ws.max_row - 1  # exclude header
        print(f"📊 Appending to existing file: {filepath} ({existing_rows} existing entries)")

    return load_workbook(filepath)


# ─── Excel Write (thread-safe) ────────────────────────────────────────────────

_excel_lock = threading.Lock()

def save_to_excel(wb, filepath, source, filename, text):
    """Append a new row to the Excel sheet."""
    with _excel_lock:
        ws   = wb.active
        row  = ws.max_row + 1
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        ws.append([row - 1, ts, source, filename, text])

        # Wrap text in the extracted text column
        ws.cell(row=row, column=5).alignment = Alignment(wrap_text=True)

        # Zebra striping
        if row % 2 == 0:
            fill = PatternFill("solid", fgColor="DEEAF1")
            for col in range(1, 6):
                ws.cell(row=row, column=col).fill = fill

        wb.save(filepath)
        print(f"✅ [{ts}] Saved → {source} | {filename} | {len(text)} chars")


# ─── OCR ─────────────────────────────────────────────────────────────────────

def run_ocr(img):
    """Extract text from a PIL Image."""
    try:
        text = pytesseract.image_to_string(img, lang="eng").strip()
        return text if text else "[No text detected]"
    except Exception as e:
        return f"[OCR error: {e}]"


# ─── Clipboard Monitor ────────────────────────────────────────────────────────

def monitor_clipboard(wb, filepath):
    """Watch clipboard for new screenshot images (PrtScn key)."""
    last_hash = None
    print("📋 Watching clipboard (PrtScn)...")

    while True:
        try:
            img = ImageGrab.grabclipboard()
            if isinstance(img, Image.Image):
                img_hash = hashlib.md5(img.tobytes()).hexdigest()
                if img_hash != last_hash:
                    last_hash = img_hash
                    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"clipboard_{ts}.png"
                    text     = run_ocr(img)
                    save_to_excel(wb, filepath, "Clipboard (PrtScn)", filename, text)
        except Exception:
            pass
        time.sleep(1)


# ─── Folder Watcher ───────────────────────────────────────────────────────────

class ScreenshotHandler(FileSystemEventHandler):
    def __init__(self, wb, filepath):
        self.wb       = wb
        self.filepath = filepath

    def on_created(self, event):
        if event.is_directory:
            return
        path = event.src_path
        if not path.lower().endswith((".png", ".jpg", ".jpeg", ".bmp")):
            return
        time.sleep(0.8)  # wait for OS to finish writing the file
        try:
            img      = Image.open(path)
            text     = run_ocr(img)
            filename = os.path.basename(path)
            save_to_excel(self.wb, self.filepath, "Screenshots Folder", filename, text)
        except Exception as e:
            print(f"⚠️  Could not process {path}: {e}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Screenshot → Excel (OCR)")
    parser.add_argument("--output", default=DEFAULT_XLS,        help="Output Excel file path")
    parser.add_argument("--folder", default=DEFAULT_SHOTS_DIR,  help="Screenshots folder to watch")
    args = parser.parse_args()

    print("\n🖥️  Screenshot → Excel Monitor")
    print("=" * 45)

    # Check Tesseract
    try:
        pytesseract.get_tesseract_version()
        print("✅ Tesseract OCR detected")
    except Exception:
        print("❌ Tesseract not found!")
        print("   Download from: https://github.com/UB-Mannheim/tesseract/wiki")
        print(f"   Install to: {TESSERACT_PATH}")
        sys.exit(1)

    wb       = init_excel(args.output)
    observer = None

    # Start clipboard monitor thread
    clip_thread = threading.Thread(
        target=monitor_clipboard,
        args=(wb, args.output),
        daemon=True
    )
    clip_thread.start()

    # Start folder watcher
    if os.path.exists(args.folder):
        handler  = ScreenshotHandler(wb, args.output)
        observer = Observer()
        observer.schedule(handler, args.folder, recursive=False)
        observer.start()
        print(f"📁 Watching folder: {args.folder}")
    else:
        print(f"⚠️  Screenshots folder not found: {args.folder}")
        print("    Only clipboard monitoring is active.")

    print(f"📊 Saving to: {os.path.abspath(args.output)}")
    print("\nPress PrtScn or Win+PrtScn to capture. Press Ctrl+C to stop.\n")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n🛑 Stopped.")
        if observer:
            observer.stop()
            observer.join()


if __name__ == "__main__":
    main()
