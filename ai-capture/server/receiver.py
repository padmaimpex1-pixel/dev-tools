"""
receiver.py
-----------
Local HTTP server that receives AI responses from the browser extension
AND runs a screenshot OCR fallback loop for non-browser AI tools (Ollama etc.)
Saves everything to D:\screenshots_data.xlsx → AI Outputs sheet.

Usage:
    python server/receiver.py
    python server/receiver.py --output D:\my_data.xlsx --port 5555
"""

import os
import sys
import time
import json
import hashlib
import argparse
import threading
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler

# Add screenshot-to-xls to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "screenshot-to-xls"))
from screenshot_to_xls import (
    init_excel, save_ai_output, _excel_lock,
    get_active_window_title, detect_ai_tool, run_ocr,
    load_workbook, PatternFill, Alignment
)

from PIL import ImageGrab, Image
import pytesseract

DEFAULT_XLS  = r"D:\screenshots_data.xlsx"
DEFAULT_PORT = 5555

# Track total saved count for popup status
_saved_count = 0
_saved_hashes = set()

# ─── HTTP Server (receives from browser extension) ────────────────────────────

class AIOutputHandler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        pass  # suppress default access log

    def do_GET(self):
        if self.path == "/status":
            self._json({"status": "ok", "saved": _saved_count})
        else:
            self._json({"status": "ok"})

    def do_POST(self):
        global _saved_count
        if self.path == "/ai-output":
            length  = int(self.headers.get("Content-Length", 0))
            body    = self.rfile.read(length)
            try:
                data = json.loads(body)
                text = data.get("text", "").strip()
                tool = data.get("tool", "Unknown")
                title= data.get("page_title", "")
                url  = data.get("url", "")

                h = hashlib.md5(text.encode("utf-8", errors="ignore")).hexdigest()
                if h not in _saved_hashes and len(text) > 30:
                    _saved_hashes.add(h)
                    save_ai_output(self.server.wb, self.server.filepath, tool, f"{title} | {url}", text)
                    _saved_count += 1

                self._json({"status": "saved"})
            except Exception as e:
                self._json({"status": "error", "msg": str(e)})
        else:
            self._json({"status": "unknown endpoint"})

    def do_OPTIONS(self):
        # CORS preflight for browser extension
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def _json(self, data):
        body = json.dumps(data).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self._cors_headers()
        self.end_headers()
        self.wfile.write(body)

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")


# ─── Screenshot OCR fallback (for Ollama, desktop AI apps) ───────────────────

TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# Only OCR-poll these tools (browser tools are covered by extension)
OCR_FALLBACK_TOOLS = ["Ollama", "LM Studio", "AnythingLLM", "Jan ", "GPT4All"]

def ocr_poll_loop(wb, filepath, interval=3):
    """
    Every `interval` seconds: if an OCR-fallback AI tool window is focused,
    screenshot it, OCR it, and save if content has changed.
    """
    global _saved_count
    last_hash = None
    print(f"OCR fallback active for: {', '.join(OCR_FALLBACK_TOOLS)}")

    while True:
        try:
            window = get_active_window_title()
            is_ocr_tool = any(t.lower() in window.lower() for t in OCR_FALLBACK_TOOLS)

            if is_ocr_tool:
                img  = ImageGrab.grab()
                text = run_ocr(img)
                h    = hashlib.md5(text.encode("utf-8", errors="ignore")).hexdigest()

                if h != last_hash and len(text.strip()) > 30:
                    last_hash = h
                    tool = next(
                        (t for t in OCR_FALLBACK_TOOLS if t.lower() in window.lower()),
                        window
                    )
                    save_ai_output(wb, filepath, tool, window, text)
                    _saved_count += 1
        except Exception:
            pass
        time.sleep(interval)


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="AI Output Capture Server")
    parser.add_argument("--output", default=DEFAULT_XLS,  help="Excel output path")
    parser.add_argument("--port",   default=DEFAULT_PORT, type=int, help="Server port")
    args = parser.parse_args()

    wb = init_excel(args.output)

    print("\n AI Output Capture Server")
    print("=" * 45)
    print(f"  HTTP server  : http://localhost:{args.port}")
    print(f"  Excel output : {args.output}")
    print(f"  Extension URL: http://localhost:{args.port}/ai-output")
    print("\n  Sources:")
    print("  [Extension] Claude, ChatGPT, Perplexity, Gemini, Copilot, Grok, DeepSeek, Mistral")
    print("  [OCR Loop ] Ollama, LM Studio, AnythingLLM, Jan, GPT4All")
    print("\nPress Ctrl+C to stop.\n")

    # Start OCR fallback thread
    ocr_thread = threading.Thread(target=ocr_poll_loop, args=(wb, args.output), daemon=True)
    ocr_thread.start()

    # Start HTTP server
    server = HTTPServer(("localhost", args.port), AIOutputHandler)
    server.wb       = wb
    server.filepath = args.output
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
