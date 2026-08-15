# -*- coding: utf-8 -*-
"""快速探针：打开粉丝性别和达人性别下拉框，dump真实选项文案 + 监听所有请求"""
import sys, json, time
from pathlib import Path
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except:
    pass
from playwright.sync_api import sync_playwright

DEBUG = Path(__file__).parent / "debug"
all_requests = []

def log(m): print(m, flush=True)

pw = sync_playwright().start()
browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
ctx = browser.contexts[0]

page = None
for p in ctx.pages:
    if "kwaixiaodian" in p.url or "daren" in p.url:
        page = p
        break
if not page:
    page = ctx.new_page()
    page.goto("https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro", timeout=60000, wait_until="domcontentloaded")
    time.sleep(12)
page.bring_to_front()
log(f"页面: {page.url[:80]}")

# 监听所有请求（不仅 promoter/list）
def on_req(req):
    if req.method == "POST" and ("list" in req.url or "promoter" in req.url or "filter" in req.url or "search" in req.url):
        all_requests.append({"url": req.url, "body": req.post_data})

page.on("request", on_req)

# ===== 粉丝性别 =====
log("\n=== 粉丝性别 ===")
page.reload(wait_until="domcontentloaded")
time.sleep(5)
try:
    page.locator("text=粉丝性别").first.click(timeout=8000)
    time.sleep(2)
    vis = page.evaluate("""() => {
        const els = document.querySelectorAll('li, span, div, label, p');
        const texts = [];
        for (const el of els) {
            const t = el.textContent.trim();
            if (t.length > 0 && t.length < 20) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) texts.push(t);
            }
        }
        return [...new Set(texts)].slice(0, 80);
    }""")
    log(f"可见文本(去重): {vis}")
    (DEBUG / "probe_gender_fan.txt").write_text("\n".join(vis), encoding="utf-8")
except Exception as e:
    log(f"打开粉丝性别失败: {e}")

# 点击第一个选项看请求
for opt in vis:
    if any(k in opt for k in ["女", "男", "均", "女多", "男多"]):
        log(f"尝试点击: {opt}")
        before = len(all_requests)
        try:
            page.locator(f"text={opt}").first.click(timeout=5000)
            time.sleep(3)
            after = len(all_requests)
            if after > before:
                log(f"  → 新请求! url={all_requests[-1]['url'][-60:]}")
                log(f"  → body={all_requests[-1]['body'][:300]}")
            else:
                log(f"  → 无新请求")
        except Exception as e2:
            log(f"  点击失败: {e2}")
        break

# ===== 达人性别 =====
log("\n=== 达人性别 ===")
page.reload(wait_until="domcontentloaded")
time.sleep(5)
try:
    page.locator("text=达人性别").first.click(timeout=8000)
    time.sleep(2)
    vis2 = page.evaluate("""() => {
        const els = document.querySelectorAll('li, span, div, label, p, button');
        const texts = [];
        for (const el of els) {
            const t = el.textContent.trim();
            if (t.length > 0 && t.length < 20) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) texts.push(t);
            }
        }
        return [...new Set(texts)].slice(0, 80);
    }""")
    log(f"可见文本(去重): {vis2}")
    (DEBUG / "probe_gender_seller.txt").write_text("\n".join(vis2), encoding="utf-8")
except Exception as e:
    log(f"打开达人性别失败: {e}")

for opt in vis2:
    if opt in ["男", "女"] or "男" in opt or "女" in opt:
        log(f"尝试点击: {opt}")
        before = len(all_requests)
        try:
            page.locator(f"text={opt}").first.click(timeout=5000)
            time.sleep(3)
            after = len(all_requests)
            if after > before:
                log(f"  → 新请求! url={all_requests[-1]['url'][-60:]}")
                log(f"  → body={all_requests[-1]['body'][:500]}")
            else:
                log(f"  → 无新请求")
        except Exception as e2:
            log(f"  点击失败: {e2}")
        break

# dump 所有请求
log(f"\n总共捕获 {len(all_requests)} 个请求")
for i, r in enumerate(all_requests):
    log(f"  [{i}] {r['url'][-80:]} | body={r['body'][:200]}")

try: browser.close()
except: pass
log("DONE")
