# -*- coding: utf-8 -*-
"""截图+DOM探针：点击粉丝性别后看实际渲染了什么"""
import sys, time
from pathlib import Path
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except:
    pass
from playwright.sync_api import sync_playwright

DEBUG = Path(__file__).parent / "debug"
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
page.reload(wait_until="domcontentloaded")
time.sleep(6)

# 截图：点击前
page.screenshot(path=str(DEBUG / "before_click_fan_gender.png"))
print("截图1: 点击前")

# 尝试点击"粉丝性别"
print("\n尝试定位'粉丝性别'元素...")
locs = page.locator("text=粉丝性别").all()
print(f"  找到 {len(locs)} 个匹配元素")
for i, loc in enumerate(locs):
    try:
        box = loc.bounding_box()
        tag = loc.evaluate("el => el.tagName + ' / ' + el.className")
        vis = loc.is_visible()
        print(f"  [{i}] tag={tag} visible={vis} box={box}")
    except:
        print(f"  [{i}] 无法获取信息")

# 点击第一个可见的
for i, loc in enumerate(locs):
    try:
        if loc.is_visible():
            print(f"\n点击元素 [{i}]...")
            loc.click(timeout=5000)
            time.sleep(3)
            break
    except:
        pass

# 截图：点击后
page.screenshot(path=str(DEBUG / "after_click_fan_gender.png"))
print("截图2: 点击后")

# dump 新出现的可见元素
new_vis = page.evaluate("""() => {
    const els = document.querySelectorAll('*');
    const result = [];
    for (const el of els) {
        const t = el.textContent.trim();
        const rect = el.getBoundingClientRect();
        if (t.length > 0 && t.length < 30 && rect.width > 0 && rect.height > 0) {
            if (t.includes('女') || t.includes('男') || t.includes('均') || t.includes('多')) {
                result.push({
                    tag: el.tagName,
                    cls: (el.className || '').toString().substring(0, 50),
                    text: t,
                    x: Math.round(rect.x),
                    y: Math.round(rect.y),
                    w: Math.round(rect.width),
                    h: Math.round(rect.height)
                });
            }
        }
    }
    return result;
}""")
print(f"\n含性别关键词的可见元素 ({len(new_vis)} 个):")
for e in new_vis:
    print(f"  <{e['tag']} class='{e['cls']}'> '{e['text']}' @({e['x']},{e['y']}) {e['w']}x{e['h']}")

try: browser.close()
except: pass
print("\nDONE")
