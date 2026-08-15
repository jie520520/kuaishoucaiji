#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""打开达人详情页，定位页面上的'分销销售额'文案及其对应数值，
并与 promoter/list 的 distPayOrderAmt30d 等字段交叉验证。"""
import json, time, re
from pathlib import Path
from playwright.sync_api import sync_playwright

DEBUG = Path("c:/Users/Administrator/Desktop/颜阿娇/快手达人采集工具/debug")
LOG = DEBUG / "probe_dist_log.txt"
TXT = DEBUG / "probe_dist_pagetext.txt"
OUT = DEBUG / "probe_dist.json"

def log(m):
    s = f"{time.strftime('%H:%M:%S')} {m}"
    print(s)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(s + "\n")

def main():
    LOG.write_text("", encoding="utf-8")
    pid = "862635530"
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = ctx.new_page()
        urls = [
            f"https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro/daren-detail?promoterId={pid}",
            f"https://cps.kwaixiaodian.com/zone/daren-match/daren-detail?promoterId={pid}",
        ]
        ok = False
        for u in urls:
            log(f"GOTO {u}")
            try:
                page.goto(u, wait_until="load", timeout=25000)
            except Exception as e:
                log(f"ERR {e}")
            try: page.wait_for_load_state("networkidle", timeout=12000)
            except Exception: pass
            time.sleep(5)
            log(f"URL={page.url} TITLE={page.title()}")
            if "login" not in page.url.lower():
                ok = True
                break
        if not ok:
            log("!!! login redirect")
            with open(OUT,"w",encoding="utf-8") as f: json.dump({"error":"login"},f)
            return

        # 滚动触发懒加载
        for y in (400, 900, 1500, 2200):
            page.evaluate(f"window.scrollTo(0,{y})")
            time.sleep(2)

        text = page.evaluate("document.body.innerText")
        with open(TXT, "w", encoding="utf-8") as f:
            f.write(text)
        log(f"PAGE_TEXT len={len(text)}")

        # 在文本中查找 分销销售额 上下文
        lines = text.splitlines()
        ctx_lines = []
        for i, ln in enumerate(lines):
            if "分销" in ln or "销售额" in ln or "GMV" in ln or "近30" in ln or "30日" in ln:
                a = max(0, i-2); b = min(len(lines), i+3)
                ctx_lines.append("\n".join(lines[a:b]))
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump({"matched_blocks": ctx_lines}, f, ensure_ascii=False, indent=2)
        log(f"MATCHED {len(ctx_lines)} blocks -> {OUT}")
        # 打印匹配块
        for blk in ctx_lines[:25]:
            print("----")
            print(blk)
        try: browser.close()
        except Exception: pass

if __name__ == "__main__":
    main()
