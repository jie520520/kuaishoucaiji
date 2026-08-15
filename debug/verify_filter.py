#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用真实列表数据验证分销销售额过滤：取多页，解析 PromoterLiveAvgGMV30d，
统计 <5000 会被剔除的比例（确认过滤有意义）。"""
import json, time
from pathlib import Path
from playwright.sync_api import sync_playwright

DEBUG = Path("c:/Users/Administrator/Desktop/颜阿娇/快手达人采集工具/debug")
OUT = DEBUG / "verify_filter.json"
THR = 5000

with sync_playwright() as pw:
    browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
    ctx = browser.contexts[0] if browser.contexts else browser.new_context()
    page = ctx.new_page()
    page.goto("https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro",
              wait_until="load", timeout=30000)
    try: page.wait_for_load_state("networkidle", timeout=12000)
    except Exception: pass
    time.sleep(4)

    rows = []
    for off in range(0, 100, 10):
        body = {"orderField":0,"orderType":1,"limit":10,"offset":off,"type":0,
                "contentTag":["755"],"hotSaleChannelIdList":[],"hotSaleSubChannelId":[],"hasContact":1}
        raw = page.evaluate("""async ({url, body}) => {
            const r = await fetch(url,{method:'POST',headers:{'Content-Type':'application/json','Accept':'application/json'},body:JSON.stringify(body),credentials:'include'});
            return await r.text();
        }""", {"url":"https://cps.kwaixiaodian.com/distribute/pc/seller/promoter/list","body":body})
        d = json.loads(raw)
        for it in d.get("data",{}).get("promoterList",[]):
            vi = it.get("viewInfo") or {}
            try: val = float(vi.get("PromoterLiveAvgGMV30d") or "")
            except (ValueError, TypeError): val = None
            rows.append({"pid": it.get("promoterId"), "nick": it.get("nickname"),
                         "场均销售额": val})
        time.sleep(0.8)

    total = len(rows)
    unknown = sum(1 for r in rows if r["场均销售额"] is None)
    drop = sum(1 for r in rows if r["场均销售额"] is not None and r["场均销售额"] < THR)
    keep = total - drop - unknown
    print(f"总样本 {total} | 未知 {unknown} | 剔除(<{THR}) {drop} | 保留 {keep}")
    print("--- 将被剔除的示例 ---")
    for r in rows:
        v = r["场均销售额"]
        if v is not None and v < THR:
            print(f"  {r['nick']} = {v}")
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"total":total,"unknown":unknown,"drop":drop,"keep":keep,"rows":rows},
                  f, ensure_ascii=False, indent=2)
    try: browser.close()
    except Exception: pass
