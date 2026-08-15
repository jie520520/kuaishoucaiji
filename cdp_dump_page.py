# -*- coding: utf-8 -*-
"""Dump 页面完整文本，确认筛选维度真实文案"""
import sys, time
from pathlib import Path
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
from playwright.sync_api import sync_playwright

DEBUG = Path(__file__).parent / "debug"
browser = None
try:
    pw = sync_playwright().start()
    browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
    ctx = browser.contexts[0]
    # 找达人广场
    page = None
    for p in ctx.pages:
        if "kwaixiaodian" in p.url or "daren" in p.url:
            page = p
            break
    if not page:
        page = ctx.new_page()
        page.goto("https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro", timeout=60000, wait_until="domcontentloaded")
        time.sleep(12)
    else:
        print(f"找到页面: {page.url[:100]}")
        page.bring_to_front()
        time.sleep(3)
    
    # 等待更久确保加载
    time.sleep(8)
    print(f"URL: {page.url}")
    print(f"Title: {page.title()}")
    
    # dump 页面文本
    text = page.evaluate("document.body.innerText")
    (DEBUG / "cdp_page_text.txt").write_text(text, encoding="utf-8")
    print(f"页面文本已保存 ({len(text)} 字符)")
    
    # 尝试找筛选面板元素
    filters = page.evaluate("""() => {
        const els = document.querySelectorAll('*');
        const result = [];
        for (const el of els) {
            const t = el.textContent.trim();
            if (t.length > 0 && t.length < 30 && (
                t.includes('粉丝') || t.includes('达人') || t.includes('年龄') ||
                t.includes('性别') || t.includes('城市') || t.includes('地域') ||
                t.includes('联系') || t.includes('坑位')
            )) {
                const tag = el.tagName;
                const cls = el.className || '';
                result.push({tag, cls, text: t});
            }
        }
        return result;
    }""")
    print(f"\n筛选相关元素 ({len(filters)} 个):")
    for f in filters[:40]:
        print(f"  <{f['tag']} class='{f['cls'][:40]}'> {f['text']}")
    
except Exception as e:
    print(f"ERROR: {e}")
finally:
    if browser:
        try: browser.close()
        except: pass
