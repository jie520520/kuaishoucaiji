#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通过 CDP 接管已登录 Chrome，抓取 promoter/list 与 promoter/info 响应，
定位"分销销售额/近30日"相关字段名。"""
import json
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

DEBUG = Path("c:/Users/Administrator/Desktop/颜阿娇/快手达人采集工具/debug")
DEBUG.mkdir(parents=True, exist_ok=True)
OUT = DEBUG / "capture_sales.json"
LOG = DEBUG / "capture_sales_log.txt"

def log(m):
    s = f"{time.strftime('%H:%M:%S')} {m}"
    print(s)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(s + "\n")

def main():
    LOG.write_text("", encoding="utf-8")
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = ctx.new_page()

        url = "https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro"
        log(f"GOTO {url}")
        try:
            page.goto(url, wait_until="load", timeout=30000)
        except Exception as e:
            log(f"GOTO_ERR {e}")
        # 等待加载与可能的跳转
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        time.sleep(5)
        log(f"URL={page.url}")
        log(f"TITLE={page.title()}")

        if "login" in page.url.lower():
            log("!!! 跳转到登录页，未登录，停止")
            with open(OUT, "w", encoding="utf-8") as f:
                json.dump({"error": "login_redirect", "url": page.url}, f, ensure_ascii=False, indent=2)
            try: browser.close()
            except Exception: pass
            return

        # 1) 抓 promoter/list 一页样本
        list_body = {
            "orderField": 0, "orderType": 1, "limit": 10, "offset": 0,
            "type": 0, "contentTag": ["755"], "hotSaleChannelIdList": [],
            "hotSaleSubChannelId": [], "hasContact": 1,
        }
        log("FETCH promoter/list (sample)")
        data = {}
        for attempt in range(3):
            try:
                raw = page.evaluate(
                    """async ({url, body}) => {
                        const r = await fetch(url, {
                            method:'POST',
                            headers:{'Content-Type':'application/json','Accept':'application/json'},
                            body: JSON.stringify(body), credentials:'include'
                        });
                        return await r.text();
                    }""",
                    {"url": "https://cps.kwaixiaodian.com/distribute/pc/seller/promoter/list",
                     "body": list_body},
                )
                data = json.loads(raw)
                log(f"LIST result={data.get('result')} error={data.get('error_msg','')}")
                break
            except Exception as e:
                log(f"LIST_ERR (try {attempt}) {e}")
                time.sleep(2)

        plist = data.get("data", {}).get("promoterList", []) if isinstance(data, dict) else []
        log(f"LIST got {len(plist)} items")
        info_data = {}
        if plist:
            pid = plist[0].get("promoterId")
            log(f"FETCH promoter/info pid={pid}")
            for attempt in range(3):
                try:
                    raw2 = page.evaluate(
                        """async ({url, pid}) => {
                            const r = await fetch(url + '?promoterId=' + pid + '&type=1', {
                                headers:{'Accept':'application/json'}, credentials:'include'
                            });
                            return await r.text();
                        }""",
                        {"url": "https://cps.kwaixiaodian.com/distribute/pc/seller/promoter/info",
                         "pid": pid},
                    )
                    info_data = json.loads(raw2)
                    log(f"INFO result={info_data.get('result')} error={info_data.get('error_msg','')}")
                    break
                except Exception as e:
                    log(f"INFO_ERR (try {attempt}) {e}")
                    time.sleep(2)

        out = {
            "list_result": data.get("result") if isinstance(data, dict) else None,
            "list_first_item": plist[0] if plist else None,
            "info_result": info_data.get("result") if isinstance(info_data, dict) else None,
            "info_data": info_data.get("data") if isinstance(info_data, dict) else None,
        }
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        log(f"SAVED {OUT}")
        try: browser.close()
        except Exception: pass

if __name__ == "__main__":
    main()
