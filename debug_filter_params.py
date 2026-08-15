# -*- coding: utf-8 -*-
"""调试脚本3：全维度筛选参数采集 —— 抓取各筛选维度"选项文本 → API参数值"映射

已知（fan_age_all_requests.txt 实证）：
  粉丝年龄点击选项后，promoter/list 请求体出现 fansAgeFeature: [数字ID]

本脚本自动完成：
  对每个维度（达人性别/粉丝年龄/粉丝性别/粉丝城市划分/达人地域）：
    1. 点击维度名展开浮层
    2. body diff 提取浮层选项文本（真实文案）
    3. 逐个点击选项，记录每次点击后的 promoter/list 请求体
  输出：
    debug/filter_params_dims.txt     各维度真实选项文案
    debug/filter_params_requests.txt 点击过程中的全部请求体（时间顺序+点击标注）
    debug/filter_params_log.txt      完整操作日志

注意：点击选项后浮层可能收起，每个选项点击前会重新展开维度。
"""
import sys
import os
import re
import json
import time
import difflib
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

DIMS = ["达人性别", "粉丝年龄", "粉丝性别", "粉丝城市划分", "达人地域"]

captured = []          # 按时间顺序的请求记录
log_lines = []         # 操作日志


def log(msg):
    print(msg, flush=True)
    log_lines.append(str(msg))


def dump_to(name, text):
    path = DEBUG_DIR / name
    path.write_text(text, encoding="utf-8")
    log(f"  💾 已保存: {path.name}")


def on_request(request):
    try:
        if "promoter/list" in request.url and request.method == "POST":
            body = json.loads(request.post_data or "{}")
            captured.append({"t": time.time(), "body": body, "note": ""})
    except Exception:
        pass


def mark_last(note):
    if captured:
        captured[-1]["note"] = note


def click_dimension(page, dim):
    try:
        loc = page.get_by_text(dim, exact=True).first
        loc.scroll_into_view_if_needed(timeout=4000)
        time.sleep(0.3)
        loc.click(timeout=4000)
        time.sleep(1.5)
        return True
    except Exception as e:
        log(f"  ⚠️ 展开维度[{dim}]失败: {e}")
        return False


def visible_texts_added(page, body_before):
    """点击维度后，body 相对 before 的新增可见文本行"""
    try:
        body_after = page.locator("body").inner_text(timeout=5000)
        diff = difflib.ndiff(body_before.splitlines(), body_after.splitlines())
        added = []
        for line in diff:
            if line.startswith("+ "):
                t = line[2:].strip()
                if t and not t.isdigit():
                    added.append(t)
        return added, body_after
    except Exception:
        return [], body_before


def click_option(page, dim, opt):
    """点击浮层选项，先确保浮层展开"""
    for _ in range(2):
        try:
            loc = page.get_by_text(opt, exact=True).first
            if loc.count() > 0 and loc.is_visible():
                loc.click(timeout=3000)
                time.sleep(1.6)  # 等列表请求发出
                return True
        except Exception:
            pass
        # 没找到/不可见 → 重新展开维度
        click_dimension(page, dim)
        time.sleep(0.8)
    return False


def main():
    log("🚀 启动浏览器（复用登录态）...")
    pw = sync_playwright().start()
    context = pw.chromium.launch_persistent_context(
        user_data_dir=str(USER_DATA),
        headless=False,
        args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        viewport={"width": 1440, "height": 900},
        locale="zh-CN",
    )
    page = context.new_page()
    page.on("request", on_request)

    log("📡 导航至达人广场...")
    page.goto(DAREN_URL, timeout=60000, wait_until="domcontentloaded")
    time.sleep(5)

    if "login" in page.url.lower():
        log("⚠️ 需要登录，请扫码（最多等120秒）...")
        for i in range(60):
            time.sleep(2)
            if "login" not in page.url.lower():
                log("✅ 登录成功")
                break
        time.sleep(4)

    # 页面初始请求作为基线
    captured.clear()
    time.sleep(1)

    dims_report = []
    try:
        body_base = page.locator("body").inner_text(timeout=5000)
    except Exception:
        body_base = ""

    for dim in DIMS:
        log(f"\n===== 维度: {dim} =====")
        if not click_dimension(page, dim):
            continue
        added, _ = visible_texts_added(page, body_base)
        opts = []
        for t in added:
            if len(t) <= 30 and t != dim and "快手号" not in t and "\n" not in t:
                opts.append(t)
        log(f"  展开后新增可见文本: {opts[:25]}")
        dims_report.append(f"===== {dim} =====\n" + "\n".join(opts))

        for opt in opts[:8]:
            if opt == dim:
                continue
            mark_last(f"[{dim}] 点击前")
            ok = click_option(page, dim, opt)
            if ok:
                log(f"  ✅ 点击: {opt}")
                mark_last(f"[{dim}] 点击: {opt}")
            else:
                log(f"  ❌ 点不到: {opt}")

        try:
            page.keyboard.press("Escape")
            time.sleep(0.5)
        except Exception:
            pass

    dump_to("filter_params_dims.txt", "\n\n".join(dims_report))

    req_lines = [f"共 {len(captured)} 个 promoter/list 请求：", "=" * 60]
    for i, rec in enumerate(captured):
        req_lines.append(f"#{i} {rec['note']}  @{time.strftime('%H:%M:%S', time.localtime(rec['t']))}")
        req_lines.append(json.dumps(rec["body"], ensure_ascii=False, indent=1))
        req_lines.append("-" * 50)
    dump_to("filter_params_requests.txt", "\n".join(req_lines))
    dump_to("filter_params_log.txt", "\n".join(log_lines))

    log("\n🏁 采集完成！结果在 debug/filter_params_*.txt。按任意键关闭浏览器...")
    try:
        input()
    except Exception:
        time.sleep(3)
    context.close()
    pw.stop()


if __name__ == "__main__":
    main()
