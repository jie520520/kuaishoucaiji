# -*- coding: utf-8 -*-
"""
CDP 接管用户已登录 Chrome，自动点击筛选维度抓取真实请求参数。
用法：先以 --remote-debugging-port=9222 启动用户 Chrome，再运行本脚本。
"""
import sys
import json
import time
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
DEBUG_DIR = BASE_DIR / "debug"
DEBUG_DIR.mkdir(exist_ok=True)
DAREN_URL = "https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro"

captured = []  # 所有 promoter/list 请求体（按时间顺序）
RESULT = {}    # dim -> {option -> 相关字段}

def log(msg):
    print(msg, flush=True)

def extract_filter_fields(body):
    """从请求体提取所有与筛选相关的字段"""
    keys = ["fansAgeFeature", "fansGenderFeature", "fansCityFeature", "fansCityLevel",
            "gender", "sellerGender", "regionCode", "provinceId", "cityId", "cityCode",
            "areaCode", "province", "hasContact", "noSlotFee", "recruiting"]
    out = {}
    for k in keys:
        if k in body:
            out[k] = body[k]
    return out

def main():
    log("🔌 连接用户 Chrome (CDP 9222)...")
    pw = sync_playwright().start()
    browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
    ctx = browser.contexts[0]
    log(f"  上下文页面: {[p.url[:80] for p in ctx.pages]}")

    # 找达人广场页面，没有就新开
    page = None
    for p in ctx.pages:
        if "kwaixiaodian" in p.url or "daren" in p.url:
            page = p
            break
    if page is None:
        log("  未找到达人广场页面，新开标签页...")
        page = ctx.new_page()
        page.goto(DAREN_URL, timeout=60000, wait_until="domcontentloaded")
        time.sleep(8)
    log(f"  使用页面: {page.url[:100]}")
    page.bring_to_front()

    # 拦截请求
    def on_request(request):
        if "promoter/list" in request.url and request.method == "POST":
            try:
                body = json.loads(request.post_data or "{}")
                captured.append(body)
            except Exception:
                pass
    page.on("request", on_request)

    def click_text(text, wait=2.8):
        """点击页面文本，返回是否成功"""
        try:
            page.get_by_text(text, exact=True).first.click(timeout=6000)
            time.sleep(wait)
            return True
        except Exception as e:
            log(f"  ✗ 点击「{text}」失败: {type(e).__name__}")
            return False

    def latest_filter_fields(prev_count):
        """取最新请求体中的筛选字段；无新请求返回 None"""
        if len(captured) > prev_count:
            return extract_filter_fields(captured[-1])
        return None

    # ============ 1. 粉丝年龄（6个选项） ============
    log("\n📊 ===== 1. 粉丝年龄 =====")
    age_map = {}
    for opt in ["18岁以下", "18-23岁", "24-30岁", "31-40岁", "41-50岁", "50岁以上"]:
        # 每次先刷新清空筛选
        try:
            page.reload(wait_until="domcontentloaded")
            time.sleep(5)
        except Exception:
            time.sleep(3)
        # 展开"粉丝年龄"浮层
        click_text("粉丝年龄", wait=1.5)
        before = len(captured)
        ok = click_text(opt, wait=3.5)
        if ok:
            fields = latest_filter_fields(before)
            if fields and "fansAgeFeature" in fields:
                age_map[opt] = fields["fansAgeFeature"]
                log(f"  ✅ {opt} → fansAgeFeature={fields['fansAgeFeature']}")
            else:
                log(f"  ⚠️ {opt} → fields={fields}")
        else:
            log(f"  ❌ {opt} 点击失败")
    RESULT["粉丝年龄"] = age_map
    log(f"  结果: {json.dumps(age_map, ensure_ascii=False)}")

    # 点击「收起更多」或重新加载列表，恢复初始状态
    try:
        page.reload(wait_until="domcontentloaded")
        time.sleep(6)
        log("  已刷新列表（清空筛选）")
    except Exception:
        pass

    # ============ 2. 粉丝性别 ============
    log("\n📊 ===== 2. 粉丝性别 =====")
    try:
        body = page.evaluate("document.body.innerText")
        with open(DEBUG_DIR / "cdp_fan_gender_options.txt", "w", encoding="utf-8") as f:
            f.write(body)
    except Exception:
        pass
    gender_map = {}
    for opt in ["女多男少", "男多女少", "男女均衡"]:
        try:
            page.reload(wait_until="domcontentloaded")
            time.sleep(5)
        except Exception:
            time.sleep(3)
        click_text("粉丝性别", wait=1.5)
        before = len(captured)
        ok = click_text(opt, wait=3.5)
        if ok:
            fields = latest_filter_fields(before)
            log(f"  {'✅' if fields else '⚠️'} {opt} → fields={fields}")
            if fields:
                gender_map[opt] = fields
    RESULT["粉丝性别"] = gender_map
    log(f"  结果: {json.dumps(gender_map, ensure_ascii=False)}")

    # ============ 3. 粉丝城市划分 ============
    log("\n📊 ===== 3. 粉丝城市划分 =====")
    try:
        body = page.evaluate("document.body.innerText")
        with open(DEBUG_DIR / "cdp_fan_city_options.txt", "w", encoding="utf-8") as f:
            f.write(body)
    except Exception:
        pass
    city_map = {}
    for opt in ["一线城市", "新一线城市", "二线城市", "三线城市", "四线及以下"]:
        try:
            page.reload(wait_until="domcontentloaded")
            time.sleep(5)
        except Exception:
            time.sleep(3)
        click_text("粉丝城市划分", wait=1.5)
        before = len(captured)
        ok = click_text(opt, wait=3.5)
        if ok:
            fields = latest_filter_fields(before)
            log(f"  {'✅' if fields else '⚠️'} {opt} → fields={fields}")
            if fields:
                city_map[opt] = fields
    RESULT["粉丝城市划分"] = city_map
    log(f"  结果: {json.dumps(city_map, ensure_ascii=False)}")

    # ============ 4. 达人性别 ============
    log("\n📊 ===== 4. 达人性别 =====")
    try:
        body = page.evaluate("document.body.innerText")
        with open(DEBUG_DIR / "cdp_seller_gender_options.txt", "w", encoding="utf-8") as f:
            f.write(body)
    except Exception:
        pass
    sg_map = {}
    for opt in ["男", "女"]:
        try:
            page.reload(wait_until="domcontentloaded")
            time.sleep(5)
        except Exception:
            time.sleep(3)
        click_text("达人性别", wait=1.5)
        before = len(captured)
        ok = click_text(opt, wait=3.5)
        if ok:
            fields = latest_filter_fields(before)
            log(f"  {'✅' if fields else '⚠️'} {opt} → fields={fields}")
            if fields:
                sg_map[opt] = fields
    RESULT["达人性别"] = sg_map
    log(f"  结果: {json.dumps(sg_map, ensure_ascii=False)}")

    # ============ 5. 达人地域（一个省份） ============
    log("\n📊 ===== 5. 达人地域 =====")
    try:
        body = page.evaluate("document.body.innerText")
        with open(DEBUG_DIR / "cdp_region_options.txt", "w", encoding="utf-8") as f:
            f.write(body)
    except Exception:
        pass
    region_map = {}
    for opt in ["北京", "上海", "广东"]:
        try:
            page.reload(wait_until="domcontentloaded")
            time.sleep(5)
        except Exception:
            time.sleep(3)
        click_text("达人地域", wait=1.5)
        before = len(captured)
        ok = click_text(opt, wait=3.5)
        if ok:
            fields = latest_filter_fields(before)
            log(f"  {'✅' if fields else '⚠️'} {opt} → fields={fields}")
            if fields:
                region_map[opt] = fields
    RESULT["达人地域"] = region_map
    log(f"  结果: {json.dumps(region_map, ensure_ascii=False)}")

    # ============ 保存 ============
    dump = DEBUG_DIR / "cdp_filter_mapping.json"
    dump.write_text(json.dumps(RESULT, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"\n💾 已保存: {dump}")

    # 保存所有请求体（原始）
    raw = DEBUG_DIR / "cdp_all_bodies.json"
    raw.write_text(json.dumps(captured, ensure_ascii=False, indent=1), encoding="utf-8")
    log(f"💾 原始请求体: {raw} ({len(captured)} 个)")

    log("🏁 抓取完成。浏览器保持打开。")
    # 只断开 CDP 连接，不关闭用户的浏览器
    try:
        browser.close()
    except Exception:
        pass

if __name__ == "__main__":
    main()
