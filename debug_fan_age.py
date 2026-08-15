# -*- coding: utf-8 -*-
"""调试脚本2：专攻"粉丝年龄"筛选 —— 搞清浮层交互 + 抓真实请求参数

背景：v11.5 里粉丝年龄选项 get_by_text 点不中。上次 debug_dump_filters.py
发现该维度是"浮层/弹窗控件"，但没抓到浮层展开后的选项文本，也没抓到
带年龄参数的请求体。本脚本专门解决这两点。

做法：
  1. 复用 .kuaishou_browser_data 登录态打开达人广场（已登录，无需扫码）
  2. 拦截所有请求（POST+GET），dump URL + body
  3. 点击"粉丝年龄"维度名，对比点击前后 body 全文差异 → 浮层新增文本
  4. 扫描 role=option / role=menu / role=dialog / 所有 iframe
  5. 若自动展开成功：尝试点击发现的年龄选项，抓带年龄参数的请求体
  6. 若自动展开失败：提示用户手动点选，脚本挂拦截器持续等待抓包（最多120秒）

运行前提：先停止 GUI 采集（profile 被占用会启动失败）。
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

captured = []          # 所有拦截到的请求
age_bodies = []        # 带 age 相关参数的请求体


def log(msg):
    print(msg, flush=True)


def dump_to(name, text):
    path = DEBUG_DIR / name
    path.write_text(text, encoding="utf-8")
    log(f"  💾 已保存: {path.name} ({len(text)} 字符)")


def main():
    log("🚀 启动浏览器（复用登录态，无需扫码）...")
    pw = sync_playwright().start()
    context = pw.chromium.launch_persistent_context(
        user_data_dir=str(USER_DATA),
        headless=False,
        args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        viewport={"width": 1440, "height": 900},
        locale="zh-CN",
    )
    page = context.new_page()

    # ---- 拦截所有请求 ----
    def on_request(request):
        try:
            url = request.url
            method = request.method
            body = None
            if method == "POST":
                try:
                    body = json.loads(request.post_data or "{}")
                except Exception:
                    body = request.post_data
            rec = {"method": method, "url": url, "body": body}
            captured.append(rec)
            # 只要 URL 或 body 里带 age/filter/筛选 相关字样就单拎出来
            joined = f"{url} {json.dumps(body, ensure_ascii=False)}" if body else url
            if any(k in joined.lower() for k in ["age", "filter", "promoter"]):
                age_bodies.append(rec)
                log(f"  📡 相关请求: {method} {url[:150]}")
        except Exception:
            pass

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

    # ---- 0. 先 dump 页面状态 ----
    try:
        body_before = page.locator("body").inner_text(timeout=5000)
        dump_to("fan_age_page_before.txt", body_before)
    except Exception as e:
        log(f"⚠️ before dump 失败: {e}")
        body_before = ""

    # ---- 1. 找"粉丝年龄"维度名并点击 ----
    log("🔍 查找并点击「粉丝年龄」维度...")
    try:
        loc = page.get_by_text("粉丝年龄", exact=False).first
        loc.scroll_into_view_if_needed(timeout=5000)
        time.sleep(0.5)
        loc.click(timeout=5000)
        log("✅ 已点击「粉丝年龄」")
    except Exception as e:
        log(f"❌ 点击「粉丝年龄」失败: {e}")

    time.sleep(2.0)  # 等浮层渲染

    # ---- 2. 点击后 body 全文 + diff ----
    try:
        body_after = page.locator("body").inner_text(timeout=5000)
        dump_to("fan_age_page_after.txt", body_after)
        diff = difflib.ndiff(body_before.splitlines(), body_after.splitlines())
        added = [l[2:] for l in diff if l.startswith("+ ") and l[2:].strip()]
        dump_to("fan_age_diff_added.txt", "点击「粉丝年龄」后新增文本行：\n" + "\n".join(added))
        log(f"  🔎 点击后新增文本行: {len(added)} 行")
        for line in added[:60]:
            log(f"    + {line}")
    except Exception as e:
        log(f"⚠️ diff 失败: {e}")

    # ---- 3. 扫描 role=option/menu/dialog/checkbox 元素 ----
    lines = ["===== role 元素扫描（粉丝年龄展开后） ====="]
    for role in ["option", "menuitem", "dialog", "tooltip", "combobox", "listbox"]:
        try:
            els = page.get_by_role(role).all()
            lines.append(f"--- role={role}: {len(els)} 个 ---")
            for el in els[:40]:
                try:
                    if el.is_visible():
                        t = el.inner_text(timeout=1000).strip()
                        lines.append(f"  · {t[:120]!r}")
                except Exception:
                    pass
        except Exception as e:
            lines.append(f"--- role={role} 扫描失败: {e}")
    # 含 age 关键词的元素
    try:
        els = page.get_by_text(re.compile(r"(年龄|岁|18|23|30|40|50)")).all()
        lines.append(f"--- 含年龄关键词元素: {len(els)} 个 ---")
        seen = set()
        for el in els[:60]:
            try:
                if el.is_visible():
                    t = el.inner_text(timeout=1000).strip()
                    if t and t not in seen:
                        seen.add(t)
                        lines.append(f"  · {t[:150]!r}")
            except Exception:
                pass
    except Exception as e:
        lines.append(f"--- 年龄关键词扫描失败: {e}")
    dump_to("fan_age_roles.txt", "\n".join(lines))

    # ---- 4. 扫描所有 iframe ----
    try:
        frames = page.frames
        fl = [f"iframe/frame 总数: {len(frames)}"]
        for f in frames:
            fl.append(f"  - {f.url[:200]}")
        dump_to("fan_age_frames.txt", "\n".join(fl))
    except Exception as e:
        dump_to("fan_age_frames.txt", f"frame 扫描失败: {e}")

    # ---- 5. 尝试自动点击发现的年龄选项 ----
    log("🎯 尝试自动点击年龄选项...")
    candidates = []
    try:
        opts = page.get_by_role("option").all()
        for o in opts:
            try:
                if o.is_visible():
                    candidates.append(o)
            except Exception:
                pass
    except Exception:
        pass
    clicked = False
    for o in candidates:
        try:
            t = o.inner_text(timeout=1000).strip()
            if any(k in t for k in ["岁", "年龄"]):
                o.click(timeout=3000)
                log(f"  ✅ 自动点击 role=option: {t}")
                clicked = True
                time.sleep(1.5)
                break
        except Exception:
            continue
    if not clicked:
        log("  ⚠️ 自动点不中浮层选项")

    # ---- 6. 抓包窗口：如果还没抓到带 age 的请求，提示用户手动点 ----
    def has_age_param():
        for rec in age_bodies:
            b = rec.get("body")
            if isinstance(b, dict):
                if any("age" in str(k).lower() for k in b.keys()):
                    return True
        return False

    if not has_age_param():
        log("=" * 60)
        log("🖐️ 请在浏览器里手动操作（最多等 120 秒）：")
        log("   1) 点开「粉丝年龄」筛选")
        log("   2) 随便选一个年龄区间（如 31-40岁）")
        log("   3) 选完即可，脚本会自动抓取请求参数")
        log("=" * 60)
        for i in range(60):
            time.sleep(2)
            if has_age_param():
                log("✅ 抓到带 age 参数的请求！")
                break
        time.sleep(1)

    # ---- 7. dump 所有拦截到的请求 ----
    lines = [f"共拦截 {len(captured)} 个请求，{len(age_bodies)} 个相关：", "=" * 60]
    for rec in captured:
        lines.append(f"[{rec['method']}] {rec['url'][:200]}")
        if rec["body"] is not None:
            try:
                lines.append(json.dumps(rec["body"], ensure_ascii=False, indent=1))
            except Exception:
                lines.append(str(rec["body"])[:2000])
        lines.append("-" * 50)
    dump_to("fan_age_all_requests.txt", "\n".join(lines))

    log("🏁 调试完成，结果在 debug/fan_age_*.txt。按任意键关闭浏览器...")
    try:
        input()
    except Exception:
        time.sleep(3)
    context.close()
    pw.stop()


if __name__ == "__main__":
    main()
