/**
 * content.js
 * Injected into all supported AI tool pages.
 * Uses MutationObserver to detect completed AI responses and
 * sends them to the local Python server on localhost:5555.
 */

const SERVER_URL = "http://localhost:5555/ai-output";
const DEBOUNCE_MS = 1500; // wait for streaming to finish before capturing

// ─── Tool Detection ───────────────────────────────────────────────────────────

const AI_TOOLS = {
  "claude.ai":             "Claude (Anthropic)",
  "chat.openai.com":       "ChatGPT (OpenAI)",
  "chatgpt.com":           "ChatGPT (OpenAI)",
  "perplexity.ai":         "Perplexity AI",
  "gemini.google.com":     "Google Gemini",
  "copilot.microsoft.com": "GitHub Copilot",
  "grok.com":              "Grok (xAI)",
  "chat.deepseek.com":     "DeepSeek",
  "chat.mistral.ai":       "Mistral AI",
};

function getToolName() {
  const host = window.location.hostname.replace("www.", "");
  for (const [domain, name] of Object.entries(AI_TOOLS)) {
    if (host.includes(domain)) return name;
  }
  return document.title;
}

// ─── Response Selectors (per tool) ───────────────────────────────────────────

const SELECTORS = {
  "claude.ai": [
    '[data-testid="assistant-message"]',
    '.font-claude-message',
    '.prose'
  ],
  "chat.openai.com": [
    '[data-message-author-role="assistant"]',
    '.markdown.prose'
  ],
  "chatgpt.com": [
    '[data-message-author-role="assistant"]',
    '.markdown.prose'
  ],
  "perplexity.ai": [
    '.prose',
    '[class*="answer"]',
    '[class*="response"]'
  ],
  "gemini.google.com": [
    'model-response',
    '.response-content',
    '[class*="model-response"]'
  ],
  "copilot.microsoft.com": [
    '[class*="response"]',
    '[class*="message-content"]',
    '.cib-chat-turn'
  ],
  "grok.com": [
    '[class*="message"]',
    '[class*="response"]'
  ],
  "chat.deepseek.com": [
    '[class*="ds-markdown"]',
    '[class*="message-content"]'
  ],
  "chat.mistral.ai": [
    '[class*="message"]',
    '.prose'
  ],
};

function getSelectors() {
  const host = window.location.hostname.replace("www.", "");
  for (const [domain, sels] of Object.entries(SELECTORS)) {
    if (host.includes(domain)) return sels;
  }
  return ['.prose', '[class*="response"]', '[class*="message"]', '[class*="answer"]'];
}

// ─── Capture & Send ───────────────────────────────────────────────────────────

const sentHashes = new Set();

function hashText(text) {
  let h = 0;
  for (let i = 0; i < text.length; i++) {
    h = Math.imul(31, h) + text.charCodeAt(i) | 0;
  }
  return h.toString(36);
}

function sendToServer(toolName, pageTitle, text) {
  const hash = hashText(text.trim());
  if (sentHashes.has(hash) || text.trim().length < 30) return;
  sentHashes.add(hash);

  const payload = {
    tool:       toolName,
    url:        window.location.href,
    page_title: pageTitle,
    text:       text.trim(),
    timestamp:  new Date().toISOString(),
    source:     "browser-extension"
  };

  fetch(SERVER_URL, {
    method:  "POST",
    headers: { "Content-Type": "application/json" },
    body:    JSON.stringify(payload)
  }).catch(() => {
    // Server not running — silently ignore
  });
}

function captureResponses() {
  const tool      = getToolName();
  const selectors = getSelectors();
  const title     = document.title;

  for (const sel of selectors) {
    const elements = document.querySelectorAll(sel);
    elements.forEach(el => {
      const text = el.innerText || el.textContent || "";
      if (text.trim().length > 30) {
        sendToServer(tool, title, text);
      }
    });
  }
}

// ─── MutationObserver ─────────────────────────────────────────────────────────

let debounceTimer = null;

const observer = new MutationObserver(() => {
  clearTimeout(debounceTimer);
  // Wait for streaming to stop before capturing
  debounceTimer = setTimeout(captureResponses, DEBOUNCE_MS);
});

observer.observe(document.body, {
  childList: true,
  subtree:   true,
  characterData: true
});

// Initial capture for already-loaded content
setTimeout(captureResponses, 2000);
console.log(`[AI Capture] Active on ${getToolName()} — sending to ${SERVER_URL}`);
