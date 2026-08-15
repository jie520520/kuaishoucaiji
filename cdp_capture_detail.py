#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抓取单个达人详情页请求，找出分销销售额字段API"""
import json, time, re
from pathlib import Path
from playwright.sync_api import sync_playwright

DEBUG_DIR = Path("c:/Users/Administrator/Desktop/颜阿娇/快手达人采集工具/debug")
DEBUG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = DEBUG_DIR / "capture_detail_log.txt"
REQUESTS_FILE = DEBUG_DIR / "capture_detail_requests.json"

captured = []

def log(msg):
    s = f"{time.strftime('%H:%M:%S')} {msg}"
    print(s)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(s + "\n")

def main():
    # 清空日志
    LOG_FILE.write_text("", encoding="utf-8")
    promoter_id = "733622991"
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        def handle_route(route, request):
            url = request.url
            if re.search(r"promoter/(info|detail|overview|data|analysis)|daren-detail|promoterData|overview", url):
                try:
                    body = request.post_data or ""
                except Exception:
                    body = ""
                captured.append({
                    "time": time.strftime("%H:%M:%S"),
                    "method": request.method,
                    "url": url,
                    "body": body[:2000],
                })
                log(f"REQ {request.method} {url}")
            route.continue_()

        page.route("**/*", handle_route)

        urls = [
            f"https://cps.kwaixiaodian.com/zone/daren-match/daren-detail?promoterId={promoter_id}",
            f"https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro/daren-detail?promoterId={promoter_id}",
            f"https://cps.kwaixiaodian.com/zone/daren-match/daren-detail?promoterId={promoter_id}&type=1",
        ]
        for u in urls:
            log(f"OPEN {u}")
            try:
                page.goto(u, wait_until="domcontentloaded", timeout=20000)
                time.sleep(6)
            except Exception as e:
                log(f"OPEN_ERR {e}")
            title = page.title() or ""
            log(f"TITLE {title}")
            if "暂无" not in title and "error" not in title.lower():
                break

        page.evaluate("window.scrollTo(0, 500)")
        time.sleep(4)
        page.evaluate("window.scrollTo(0, 1000)")
        time.sleep(4)

        with open(REQUESTS_FILE, "w", encoding="utf-8") as f:
            json.dump(captured, f, ensure_ascii=False, indent=2)
        log(f"SAVED {len(captured)} requests to {REQUESTS_FILE}")

        try:
            text = page.evaluate("document.body.innerText")
            with open(DEBUG_DIR / "capture_detail_page_text.txt", "w", encoding="utf-8") as f:
                f.write(text)
            log("PAGE_TEXT SAVED")
        except Exception as e:
            log(f"PAGE_TEXT_ERR {e}")

        try:
            browser.close()
        except Exception:
            pass

if __name__ == "__main__":
    main()
