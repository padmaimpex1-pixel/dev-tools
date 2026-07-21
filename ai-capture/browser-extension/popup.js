fetch("http://localhost:5555/status")
  .then(r => r.json())
  .then(data => {
    document.getElementById("status").className = "status ok";
    document.getElementById("status").textContent = `Server running — ${data.saved} responses saved`;
  })
  .catch(() => {
    document.getElementById("status").className = "status err";
    document.getElementById("status").textContent = "Server offline — run: python server/receiver.py";
  });

const tools = [
  "Claude (Anthropic)", "ChatGPT (OpenAI)", "Perplexity AI",
  "Google Gemini", "GitHub Copilot", "Grok (xAI)",
  "DeepSeek", "Mistral AI"
];
const div = document.getElementById("tools");
div.innerHTML = "<b style='font-size:12px'>Monitored tools:</b><br>" +
  tools.map(t => `<div class='tool'>✅ ${t}</div>`).join("");
