"""Pull rendered insider signal text from each HTML report."""
import os, re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

reports_dir = "static/reports"
print("Insider signal text rendered in current HTML reports:\n")
print(f"  {'Ticker':<7}  {'Signal text (first 110 chars)'}")
print("-" * 130)
for f in sorted(os.listdir(reports_dir)):
    if not f.endswith("_report.html"):
        continue
    t = f.replace("_report.html", "")
    with open(os.path.join(reports_dir, f), encoding="utf-8") as fh:
        html = fh.read()
    # Find the insider signal block (rendered around 'Insider Buying Signal' header or similar)
    m = re.search(r'insider[^<]*signal.*?<\/p>', html, re.IGNORECASE | re.DOTALL)
    if m:
        text = re.sub(r"<[^>]+>", " ", m.group(0))
        text = " ".join(text.split())
        print(f"  {t:<7}  {text[:110]}")
    else:
        m2 = re.search(r"(no open-market|net buying|net selling|no data|no insider)", html, re.IGNORECASE)
        print(f"  {t:<7}  [no insider block found]  {m2.group(0) if m2 else ''}")
