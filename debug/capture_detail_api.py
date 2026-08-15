#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""打开达人详情页，拦截所有 API 响应，定位返回'分销销售额'的接口与字段名。"""
import json, time, re
from pathlib import Path
from playwright.sync_api import sync_playwright

DEBUG = Path("c:/Users/Administrator/Desktop/颜阿娇/快手达人采集工具/debug")
LOG = DEBUG / "capture_detail_api_log.txt"
OUT = DEBUG / "capture_detail_api.json"

def log(m):
    s = f"{time.strftime('%H:%M:%S')} {m}"
    print(s)
    with open(LOG, "a", encoding="utf-8") as f: f.write(s + "\n")

def main():
    LOG.write_text("", encoding="utf-8")
    pid = "862635530"
    captures = []
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = ctx.new_page()

        def on_response(response):
            url = response.url
            ct = response.headers.get("content-type", "")
            if "json" in ct or "javascript" in ct or "text" in ct:
                if re.search(r"promoter|daren|data|analysis|overview|square", url, re.I):
                    try:
                        body = response.body()
                        txt = body.decode("utf-8", "ignore")
                    except Exception:
                        txt = ""
                    captures.append({"url": url, "len": len(txt), "snippet": txt[:500]})
                    if "分销" in txt or "fenxiao" in txt.lower() or "distPay" in txt:
                        log(f"*** MATCH {url} len={len(txt)}")

        page.on("response", on_response)
        url = f"https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro/daren-detail?promoterId={pid}"
        log(f"GOTO {url}")
        page.goto(url, wait_until="load", timeout=30000)
        try: page.wait_for_load_state("networkidle", timeout=15000)
        except Exception: pass
        # 多次滚动触发懒加载
        for y in (400, 900, 1500, 2200, 3000):
            page.evaluate(f"window.scrollTo(0,{y})")
            time.sleep(2)
        time.sleep(3)
        log(f"URL={page.url} TITLE={page.title()}")
        log(f"CAPTURED {len(captures)} responses")

        # 找出含 分销/销售额 的响应，保存完整
        hits = []
        for c in captures:
            if "分销" in c["snippet"] or "fenxiao" in c["snippet"].lower():
                hits.append(c)
        log(f"HITS {len(hits)}")
        # 保存所有响应url + 含分销的snippet
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump({
                "all_urls": [c["url"] for c in captures],
                "hits": [{"url": h["url"], "snippet": h["snippet"]} for h in hits],
            }, f, ensure_ascii=False, indent=2)
        # 打印所有 API url
        print("=== ALL API URLs ===")
        for c in captures:
            print(c["url"])
        print("=== HITS snippets ===")
        for h in hits:
            print("URL:", h["url"])
            print(h["snippet"][:800])
            print("----")
        try: browser.close()
        except Exception: pass

if __name__ == "__main__":
    main()
