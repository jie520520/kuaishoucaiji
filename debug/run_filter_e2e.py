#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""端到端验证：用真实 Chrome(CDP) 跑一个标签，确认分销销售额过滤在真实采集链路中生效。"""
import os, sys, time
sys.path.insert(0, r"c:/Users/Administrator/Desktop/颜阿娇/快手达人采集工具")
from yanajiao_scraper import YanajiaoScraper, Config
from playwright.sync_api import sync_playwright

def main():
    pw = sync_playwright().start()
    browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
    ctx = browser.contexts[0] if browser.contexts else browser.new_context()
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    if "kwaixiaodian" not in page.url:
        page.goto(Config.DAREN_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(3000)
    print("PAGE:", page.url)

    s = YanajiaoScraper()
    s.playwright = pw; s.browser = browser; s.context = ctx; s.page = page
    s._cdp_mode = True
    s.page_size = 10
    s.has_contact = True
    s.promoter_type = 0
    s.filter_config = {}
    s._page_filter_params = None
    s._stop = False
    s._on_progress = None
    s._on_tag_done = None

    tag = "健康"
    dlist = s.fetch_tag(tag)
    print(f"[1] fetch_tag('{tag}') -> {len(dlist)} 人")
    for d in dlist[:5]:
        print(f"     {d.get('昵称','?')[:14]} 场均销售额={d.get('场均销售额','')}")

    dlist = s.filter_low_dist_sales(dlist, tag)
    print(f"[2] 分销销售额过滤后 -> {len(dlist)} 人")

    s.enrich_detail_fans(dlist, tag)
    print(f"[3] enrich_detail_fans 完成")

    dlist = s.apply_filters(dlist, tag)
    print(f"[4] apply_filters 后 -> {len(dlist)} 人")

    fp = s.score_and_save(dlist, tag)
    print(f"[5] 保存: {fp}")

    s.close()
    print("DONE")

if __name__ == "__main__":
    main()
