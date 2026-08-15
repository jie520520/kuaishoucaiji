# -*- coding: utf-8 -*-
"""极简登录态探针：启动浏览器→打开达人广场→截图→判断是否登录"""
import sys
import io
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
BASE_DIR = Path(__file__).parent
USER_DATA = BASE_DIR / ".kuaishou_browser_data"
DEBUG_DIR = BASE_DIR / "debug"
DEBUG_DIR.mkdir(exist_ok=True)
LOG = DEBUG_DIR / "probe_login.txt"

def log(msg):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    print(msg, flush=True)

from playwright.sync_api import sync_playwright

def main():
    log("=== 启动登录态探针 ===")
    pw = sync_playwright().start()
    try:
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA),
            headless=False,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
            viewport={"width": 1440, "height": 900},
            locale="zh-CN",
        )
    except Exception as e:
        log(f"❌ 启动失败（可能profile被占用，请先关闭所有浏览器再试）：{e}")
        pw.stop()
        return

    page = context.new_page()
    # 收集 cookies 里是否有 kuaishou 登录相关
    cookies = context.cookies()
    ks_cookies = [c for c in cookies if "kuaishou" in c.get("domain","") or "kwai" in c.get("domain","")]
    log(f"[COOKIE] 快手相关 cookie 数量: {len(ks_cookies)}")
    for c in ks_cookies[:10]:
        log(f"  - {c.get('name')} @ {c.get('domain')}")

    url = "https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro"
    log(f"→ 打开: {url}")
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(6000)

    # 判断是否登录：看 URL 是否跳转到登录页、是否出现登录元素
    cur = page.url
    log(f"[URL] 当前URL: {cur}")
    logged_in = True
    if "login" in cur.lower() or "passport" in cur.lower() or "auth" in cur.lower():
        logged_in = False
        log("[判断] URL 跳到登录/授权页 → 未登录")

    # 截图
    shot = DEBUG_DIR / "probe_login_shot.png"
    try:
        page.screenshot(path=str(shot), full_page=False)
        log(f"[截图] 已保存: {shot}")
    except Exception as e:
        log(f"[截图] 失败: {e}")

    # 抓取页面关键文本（登录按钮/登录提示 vs 达人广场内容）
    try:
        body_text = page.evaluate("document.body ? document.body.innerText : ''")
        body_text = body_text[:3000]
        with open(DEBUG_DIR / "probe_login_body.txt", "w", encoding="utf-8") as f:
            f.write(body_text)
        log("[文本] 已保存页面正文前3000字到 probe_login_body.txt")
    except Exception as e:
        log(f"[文本] 失败: {e}")

    log("=== 探针完成，浏览器保持打开15秒供查看 ===")
    page.wait_for_timeout(15000)
    context.close()
    pw.stop()

if __name__ == "__main__":
    main()
