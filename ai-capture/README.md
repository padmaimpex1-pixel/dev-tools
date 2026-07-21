# AI Output Capture

Automatically captures **all** responses from AI tools — no manual copying needed.

## Architecture

```
Browser (Claude/ChatGPT/Perplexity/Gemini/Copilot...)
    └── browser-extension/content.js  (MutationObserver → DOM text)
            └──► POST http://localhost:5555/ai-output
                        └──► server/receiver.py
                                    └──► D:\screenshots_data.xlsx → AI Outputs sheet

Desktop AI apps (Ollama, LM Studio, Jan...)
    └── Screenshot OCR loop (every 3s when window is focused)
            └──► server/receiver.py
                        └──► D:\screenshots_data.xlsx → AI Outputs sheet
```

## Supported Tools

| Tool | Method |
|------|--------|
| Claude (Anthropic) | Browser Extension |
| ChatGPT (OpenAI) | Browser Extension |
| Perplexity AI | Browser Extension |
| Google Gemini | Browser Extension |
| GitHub Copilot | Browser Extension |
| Grok (xAI) | Browser Extension |
| DeepSeek | Browser Extension |
| Mistral AI | Browser Extension |
| Ollama (Local) | Screenshot OCR |
| LM Studio | Screenshot OCR |
| AnythingLLM | Screenshot OCR |

## Setup

### Step 1: Start the receiver server
```powershell
cd D:\GitRepos\dev-tools\ai-capture
pip install pillow pytesseract openpyxl
python server/receiver.py
```

### Step 2: Install the browser extension

**Chrome / Edge:**
1. Open `chrome://extensions` or `edge://extensions`
2. Enable **Developer mode** (top right toggle)
3. Click **Load unpacked**
4. Select the `browser-extension/` folder

**Verify:** Click the extension icon — it should show "Server running"

### Step 3: Browse AI tools
Open any supported AI tool and have a conversation. All responses are automatically saved to:
```
D:\screenshots_data.xlsx → AI Outputs sheet
```

## Excel Output

Each row in **AI Outputs** sheet contains:

| # | Timestamp | AI Tool | Window/URL | Full Response Text |
|---|-----------|---------|-----------|-------------------|

Rows are color-coded by AI tool for easy identification.

## Options

```powershell
# Custom output file
python server/receiver.py --output "C:\Users\dell\Desktop\ai_data.xlsx"

# Custom port
python server/receiver.py --port 6000
```
