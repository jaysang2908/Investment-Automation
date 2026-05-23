import json, pathlib, html, sys

# Both JSONL files: previous session + current session
JSONL_FILES = [
    pathlib.Path(r"C:\Users\justi\.claude\projects\C--Users-justi\89f4314c-561d-4a5c-9476-f9fcea833bd0.jsonl"),
    pathlib.Path(r"C:\Users\justi\.claude\projects\C--Users-justi\098f0d78-75d1-4b1d-894d-1767efc408f0.jsonl"),
]
OUT = pathlib.Path(r"C:\Users\justi\Investment-Automation\chat_export.html")

def extract_text(content):
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, dict) and block.get("type") == "text":
                parts.append(block.get("text", ""))
        return "\n".join(parts)
    return ""

rows = []
for JSONL in JSONL_FILES:
    if not JSONL.exists():
        print(f"Skipping (not found): {JSONL.name}")
        continue
    print(f"Reading {JSONL.name} ({JSONL.stat().st_size // 1024}KB)...")
    with JSONL.open(encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except json.JSONDecodeError:
                continue

            # Handle both formats
            role = obj.get("role") or obj.get("type", "")
            content_raw = obj.get("content", "")

            # queue-operation format: enqueue=user, response=assistant
            if obj.get("type") == "queue-operation":
                if obj.get("operation") == "enqueue":
                    role = "user"
                else:
                    continue
            elif obj.get("type") in ("assistant", "user"):
                role = obj["type"]
            elif role not in ("user", "assistant"):
                continue

            text = extract_text(content_raw)
            if not text.strip():
                continue
            rows.append((role, text))

print(f"Found {len(rows)} messages total")

lines_html = []
for role, text in rows:
    label = "Justin" if role == "user" else "Claude"
    bg    = "#f5f1eb" if role == "user" else "#ffffff"
    border = "#999" if role == "user" else "#c98b3c"
    escaped = html.escape(text).replace("\n", "<br>")
    lines_html.append(f"""
<div style="margin:16px 0;padding:14px 18px;background:{bg};border-left:3px solid {border};border-radius:4px;font-family:Georgia,serif;font-size:14px;color:#1a1814;line-height:1.7">
  <strong style="font-family:Arial,sans-serif;font-size:10px;letter-spacing:1.5px;text-transform:uppercase;color:#888;display:block;margin-bottom:8px">{label}</strong>
  {escaped}
</div>""")

body = "\n".join(lines_html)
html_doc = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>Chat Export — Investment Automation</title>
<style>body{{max-width:860px;margin:40px auto;padding:0 20px;background:#ece7df;font-family:Georgia,serif}}</style>
</head><body>
<h2 style="font-family:Arial,sans-serif;font-size:13px;letter-spacing:2px;text-transform:uppercase;color:#888;margin-bottom:32px">Investment Automation — Chat Export</h2>
{body}
</body></html>"""

OUT.write_text(html_doc, encoding="utf-8")
print(f"Saved: {OUT}")
