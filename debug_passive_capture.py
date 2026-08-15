# -*- coding: utf-8 -*-
"""调试脚本4：被动拦截 —— 用户手动点选筛选，抓取各维度真实参数（轻量版）

自动采集脚本被跳过（等待太久），改用被动拦截：
  1. 打开达人广场（已登录自动进入）
  2. 挂拦截器，只记录 promoter/list 请求体
  3. 提示用户在浏览器里手动点选各筛选维度
  4. 检测到足够请求后提前结束，最多等 150 秒，自动退出（无需按键）

用户操作指引（控制台也会打印）：
  ① 点开「粉丝年龄」→ 依次选「18岁以下」「31-40岁」「50岁以上」（验证ID顺序）
  ② 点开「粉丝性别」→ 选一个值
  ③ 点开「粉丝城市划分」→ 选一个值
  ④ 点开「达人性别」→ 选一个值
  ⑤ 点开「达人地域」→ 选一个省份
"""
import sys
import os
import json
import time
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
USER_DATA = BASE_DIR / ".kuaishou_browser_data"
DAREN_URL = "https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro"
DEBUG_DIR = BASE_DIR / "debug"
DEBUG_DIR.mkdir(exist_ok=True)

captured = []          # {"t": ts, "body": dict}
BASELINE = 0            # 基线请求数（进入页面时已发出的）


def log(msg):
    print(msg, flush=True)


def on_request(request):
    try:
        if "promoter/list" in request.url and request.method == "POST":
            body = json.loads(request.post_data or "{}")
            captured.append({"t": time.time(), "body": body})
    except Exception:
        pass


def main():
    log("🚀 启动浏览器（系统Chrome内核，复用登录态）...")
    pw = sync_playwright().start()
    context = pw.chromium.launch_persistent_context(
        user_data_dir=str(USER_DATA),
        channel="chrome",  # 用系统真实Chrome，避免被快手风控识别为自动化浏览器
        headless=False,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-infobars",
        ],
        viewport={"width": 1440, "height": 900},
        locale="zh-CN",
        ignore_default_args=["--enable-automation"],  # 去掉自动化标记
    )
    # 隐藏 webdriver 特征
    context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        window.chrome = window.chrome || {runtime: {}};
    """)
    page = context.new_page()
    page.on("request", on_request)

    log("📡 导航至达人广场...")
    page.goto(DAREN_URL, timeout=60000, wait_until="domcontentloaded")
    time.sleep(6)

    if "login" in page.url.lower():
        log("⚠️ 需要登录，请扫码（最多等120秒）...")
        for i in range(60):
            time.sleep(2)
            if "login" not in page.url.lower():
                log("✅ 登录成功")
                break
        time.sleep(4)

    global BASELINE
    BASELINE = len(captured)
    log(f"  基线请求 {BASELINE} 个（进入页面自动发出的，忽略）")

    log("=" * 60)
    log("🖐️ 请在浏览器里手动操作（脚本自动抓包，最多等 150 秒）：")
    log("   ① 点开「粉丝年龄」，依次点选：18岁以下、31-40岁、50岁以上")
    log("      （每点一个，列表会刷新，脚本记录对应的请求参数）")
    log("   ② 点开「粉丝性别」，点选任意一个值")
    log("   ③ 点开「粉丝城市划分」，点选任意一个值")
    log("   ④ 点开「达人性别」，点选任意一个值")
    log("   ⑤ 点开「达人地域」，点选任意一个省份")
    log("   完成后无需任何操作，脚本自动保存并退出")
    log("=" * 60)

    last_count = 0
    start = time.time()
    while time.time() - start < 150:
        time.sleep(2)
        new = len(captured) - BASELINE
        if new != last_count:
            log(f"  📦 已捕获 {new} 个筛选请求...")
            last_count = new
        if new >= 10:
            log("✅ 已捕获足够请求，自动结束")
            break

    lines = [f"基线 {BASELINE}，共捕获 {len(captured) - BASELINE} 个筛选请求：", "=" * 60]
    for i, rec in enumerate(captured[BASELINE:], 1):
        lines.append(f"#{i} @{time.strftime('%H:%M:%S', time.localtime(rec['t']))}")
        lines.append(json.dumps(rec["body"], ensure_ascii=False, indent=1))
        lines.append("-" * 50)
    dump_path = DEBUG_DIR / "filter_passive_requests_v2.txt"
    dump_path.write_text("\n".join(lines), encoding="utf-8")
    log(f"💾 已保存: {dump_path.name} ({len(lines)} 行)")

    log("🏁 完成，浏览器即将关闭。")
    context.close()
    pw.stop()


if __name__ == "__main__":
    main()
