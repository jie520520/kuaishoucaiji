#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抓取达人广场列表页文本(找'分销销售额'列名) + 取多页 distPayOrderAmt30d 数值分布，
确认分销销售额字段。"""
import json, time
from pathlib import Path
from playwright.sync_api import sync_playwright

DEBUG = Path("c:/Users/Administrator/Desktop/颜阿娇/快手达人采集工具/debug")
LOG = DEBUG / "probe_list_log.txt"
TXT = DEBUG / "probe_list_pagetext.txt"
OUT = DEBUG / "probe_list.json"

def log(m):
    s = f"{time.strftime('%H:%M:%S')} {m}"
    print(s)
    with open(LOG, "a", encoding="utf-8") as f: f.write(s + "\n")

def main():
    LOG.write_text("", encoding="utf-8")
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = ctx.new_page()
        url = "https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro"
        page.goto(url, wait_until="load", timeout=30000)
        try: page.wait_for_load_state("networkidle", timeout=12000)
        except Exception: pass
        time.sleep(5)
        text = page.evaluate("document.body.innerText")
        with open(TXT, "w", encoding="utf-8") as f: f.write(text)
        log(f"LIST PAGE TEXT len={len(text)}")
        for ln in text.splitlines():
            if any(t in ln for t in ["分销","销售额","近30","30日","GMV"]):
                print("LABEL:", ln)

        # 取多页 distPayOrderAmt30d
        samples = []
        for off in (0, 10, 20, 30):
            body = {"orderField":0,"orderType":1,"limit":10,"offset":off,"type":0,
                    "contentTag":["755"],"hotSaleChannelIdList":[],"hotSaleSubChannelId":[],"hasContact":1}
            try:
                raw = page.evaluate("""async ({url, body}) => {
                    const r = await fetch(url,{method:'POST',headers:{'Content-Type':'application/json','Accept':'application/json'},body:JSON.stringify(body),credentials:'include'});
                    return await r.text();
                }""", {"url":"https://cps.kwaixiaodian.com/distribute/pc/seller/promoter/list","body":body})
                d = json.loads(raw)
                for it in d.get("data",{}).get("promoterList",[]):
                    vi = it.get("viewInfo") or {}
                    samples.append({
                        "pid": it.get("promoterId"),
                        "nick": it.get("nickname"),
                        "distPayOrderAmt30d": vi.get("distPayOrderAmt30d"),
                        "liveAvgGmv": vi.get("liveAvgGmv"),
                        "PromoterLiveAvgGMV30d": vi.get("PromoterLiveAvgGMV30d"),
                    })
            except Exception as e:
                log(f"FETCH_ERR off={off} {e}")
            time.sleep(1)
        with open(OUT,"w",encoding="utf-8") as f: json.dump(samples,f,ensure_ascii=False,indent=2)
        log(f"SAMPLES {len(samples)} -> {OUT}")
        for s in samples:
            print(s)
        try: browser.close()
        except Exception: pass

if __name__ == "__main__":
    main()
