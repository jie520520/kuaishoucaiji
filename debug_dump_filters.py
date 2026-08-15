# -*- coding: utf-8 -*-
"""调试脚本：dump 达人广场筛选面板的真实文本结构 + 点击后拦截请求体

用途：v11.5 页面级筛选"部分失败"排查 —— 页面上的选项文本到底是什么？
做法：
  1. 复用 .kuaishou_browser_data 登录态打开达人广场
  2. dump body 全文 → debug/filters_page_dump.txt
  3. 对每个筛选维度关键词：找到元素 + 祖先容器，dump inner_text
  4. 点击维度展开，再 dump 可见选项文本
  5. 拦截 promoter/list 请求体，看快手自己用什么参数值

运行前提：先停止 GUI 采集（浏览器 profile 被占用会启动失败）。
"""
import sys
import os
import json
import time
from pathlib import Path

# 控制台输出用 UTF-8（避免 GBK 无法编码 emoji/中文）
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

captured_bodies = []


def log(msg):
    print(msg, flush=True)


def dump_to(name, text):
    path = DEBUG_DIR / name
    path.write_text(text, encoding="utf-8")
    log(f"  💾 已保存: {path.name} ({len(text)} 字符)")


def find_container_text(page, kw, max_levels=6):
    """找到含关键词的元素，向上逐层找容器，返回各层 inner_text"""
    results = []
    try:
        loc = page.get_by_text(kw, exact=False).first
        if loc.count() == 0:
            return [f"[{kw}] 未找到"]
        results.append(f"[{kw}] 元素自身 inner_text: {loc.inner_text(timeout=2000)!r}")
        cur = loc
        for level in range(1, max_levels + 1):
            cur = cur.locator("xpath=..")
            try:
                t = cur.inner_text(timeout=1500)
                results.append(f"  ↑第{level}层容器: {t[:500]!r}")
            except Exception:
                results.append(f"  ↑第{level}层: 获取失败")
                break
    except Exception as e:
        results.append(f"[{kw}] 错误: {e}")
    return results


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

    # 拦截 promoter/list 请求体
    def on_request(request):
        try:
            if "promoter/list" in request.url and request.method == "POST":
                body = json.loads(request.post_data or "{}")
                captured_bodies.append({"url": request.url, "body": body})
        except Exception:
            pass

    page.on("request", on_request)

    log("📡 导航至达人广场...")
    page.goto(DAREN_URL, timeout=60000, wait_until="domcontentloaded")
    time.sleep(5)

    # 等待登录
    if "login" in page.url.lower():
        log("⚠️ 需要扫码登录，请在浏览器中完成登录（最多等90秒）...")
        for i in range(45):
            time.sleep(2)
            if "login" not in page.url.lower():
                log("✅ 登录成功")
                break
        time.sleep(4)

    # 1. dump body 全文
    log("📄 保存页面全文...")
    try:
        body_text = page.locator("body").inner_text(timeout=5000)
        dump_to("filters_page_dump.txt", body_text)
    except Exception as e:
        log(f"⚠️ body dump 失败: {e}")
        return

    # 2. 每个筛选维度关键词：元素+祖先容器文本
    log("🔍 分析筛选维度结构...")
    dim_keywords = [
        "内容标签", "达人地域", "达人性别", "粉丝年龄",
        "粉丝性别", "粉丝城市划分", "有联系方式", "无坑位费", "招商中",
    ]
    all_lines = []
    for kw in dim_keywords:
        all_lines.append("=" * 60)
        all_lines.extend(find_container_text(page, kw))
        all_lines.append("")
    dump_to("filters_dims.txt", "\n".join(all_lines))

    # 3. 逐个点击维度展开，dump 点击后新增可见文本
    log("🖱️ 逐个展开筛选维度...")
    click_lines = []
    for kw in dim_keywords:
        try:
            loc = page.get_by_text(kw, exact=False).first
            if loc.count() == 0:
                continue
            loc.scroll_into_view_if_needed(timeout=3000)
            time.sleep(0.3)
            loc.click(timeout=3000)
            time.sleep(1.2)  # 等选项面板展开
            # dump 点击后页面上所有疑似选项的小文本（10字以内，可见）
            click_lines.append(f"----- 点击 [{kw}] 后 -----")
            try:
                all_small = page.locator(
                    "text=/^[\\u4e00-\\u9fa5A-Za-z0-9~\\-（()） ]{2,12}$/"
                ).all()
                seen = set()
                for el in all_small[:200]:
                    try:
                        if not el.is_visible():
                            continue
                        t = el.inner_text(timeout=1000).strip()
                        if t and t not in seen and len(t) <= 12:
                            seen.add(t)
                            click_lines.append(f"  · {t!r}")
                    except Exception:
                        pass
            except Exception as e:
                click_lines.append(f"  选项扫描失败: {e}")
        except Exception as e:
            click_lines.append(f"----- 点击 [{kw}] 失败: {e} -----")
        # 点击别处收起面板，避免重叠
        try:
            page.mouse.click(10, 300)
            time.sleep(0.5)
        except Exception:
            pass
    dump_to("filters_options.txt", "\n".join(click_lines))

    # 4. 手动点击几个关键选项，验证请求体参数
    log("🎯 手动点击粉丝年龄选项，验证请求参数...")
    for txt in ["31~40岁", "31-40岁", "41~50岁", "41-50岁", "50岁以上", "女性为主", "女多男少"]:
        try:
            loc = page.get_by_text(txt, exact=False).first
            if loc.count() > 0:
                loc.click(timeout=2000)
                time.sleep(1.0)
                log(f"  ✅ 点击成功: {txt}")
            else:
                log(f"  ❌ 找不到: {txt}")
        except Exception:
            log(f"  ❌ 点击失败: {txt}")

    # 5. dump 拦截到的请求体
    if captured_bodies:
        lines = []
        for c in captured_bodies:
            lines.append(json.dumps(c["body"], ensure_ascii=False, indent=1))
            lines.append("-" * 50)
        dump_to("filters_request_bodies.txt", "\n".join(lines))
        log(f"📦 拦截到 {len(captured_bodies)} 个 promoter/list 请求体")
    else:
        log("📦 未拦截到 promoter/list 请求（可能需要点击选项才触发）")

    log("🏁 调试完成，结果在 debug/ 目录。按任意键关闭浏览器...")
    try:
        input()
    except Exception:
        time.sleep(3)
    context.close()
    pw.stop()


if __name__ == "__main__":
    main()
