"""Minimal DuckDuckGo image search -> direct image URLs."""
import json
import re
import subprocess
import sys
import urllib.parse

UA = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/120.0.0.0 Safari/537.36")


def curl(url, headers=(), binary=False, out=None, timeout=45):
    cmd = ["curl", "-sSL", "--max-time", str(timeout), "-A", UA]
    for h in headers:
        cmd += ["-H", h]
    if out:
        cmd += ["-o", out, "-w", "%{http_code}"]
    cmd.append(url)
    r = subprocess.run(cmd, capture_output=True)
    return r.stdout if binary or out else r.stdout.decode("utf-8", "replace")


def search(query, limit=12):
    q = urllib.parse.quote(query)
    html = curl(f"https://duckduckgo.com/?q={q}&iax=images&ia=images")
    m = re.search(r'vqd="([0-9A-Za-z-]+)"', html) or re.search(r"vqd=([0-9A-Za-z-]+)", html)
    if not m:
        return []
    vqd = m.group(1)
    url = (f"https://duckduckgo.com/i.js?l=us-en&o=json&q={q}&vqd={vqd}"
           f"&f=,,,&p=1")
    raw = curl(url, headers=["Referer: https://duckduckgo.com/",
                             "Accept: application/json, text/javascript, */*; q=0.01",
                             "X-Requested-With: XMLHttpRequest"])
    try:
        data = json.loads(raw)
    except Exception:
        return []
    out = []
    for r in data.get("results", [])[:limit]:
        out.append({"image": r.get("image"), "w": r.get("width"), "h": r.get("height"),
                    "title": (r.get("title") or "")[:80], "src": r.get("url")})
    return out


if __name__ == "__main__":
    for r in search(" ".join(sys.argv[1:])):
        print(f"{r['w']}x{r['h']}\t{r['title']}\n\t{r['image']}")
