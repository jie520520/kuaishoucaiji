# -*- coding: utf-8 -*-
"""
CDP 接管用户已登录 Chrome，自动点击筛选维度抓取真实请求参数。
v2: 用 XPath 定位元素，更稳健。
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

captured = []
RESULT = {}

def log(msg):
    print(msg, flush=True)

def extract_filter_fields(body):
    keys = ["fansAgeFeature", "fansGenderFeature", "fansCityFeature", "fansCityLevel",
            "gender", "sellerGender", "regionCode", "provinceId", "cityId", "cityCode",
            "areaCode", "province", "hasContact", "noSlotFee", "recruiting",
            "fansAge", "fansGender", "fansCity", "sellerAge", "ageFeature"]
    out = {}
    for k in keys:
        if k in body:
            out[k] = body[k]
    # 如果没匹配到已知 key，dump 整个 body 供分析
    if not out:
        out = {"__full_body__": body}
    return out

def click_xpath(page, xpath, wait=3.0):
    """用 XPath 点击元素，返回是否成功"""
    try:
        loc = page.locator(f"xpath={xpath}").first
        loc.click(timeout=8000)
        time.sleep(wait)
        return True
    except Exception as e:
        log(f"  ✗ XPath点击失败 [{xpath[:50]}]: {type(e).__name__}")
        return False

def click_text(page, text, wait=3.0):
    """用 contains text 点击"""
    try:
        loc = page.locator(f"text={text}").first
        loc.click(timeout=8000)
        time.sleep(wait)
        return True
    except Exception:
        return False

def click_label_open(page, label_text):
    """点击筛选维度标签（展开浮层），返回是否成功"""
    # 尝试多种 XPath
    xpaths = [
        f"//*[contains(text(),'{label_text}') and not(contains(text(),'：'))]",
        f"//span[contains(text(),'{label_text}')]",
        f"//div[contains(@class,'filter') and contains(text(),'{label_text}')]",
        f"//label[contains(text(),'{label_text}')]",
    ]
    for xp in xpaths:
        if click_xpath(page, xp, wait=2.0):
            return True
    return False

def latest_filter_fields(prev_count):
    if len(captured) > prev_count:
        return extract_filter_fields(captured[-1])
    return None

def capture_dimension(page, dim_name, options, field_key=None):
    """通用：抓取某个筛选维度的各选项参数"""
    log(f"\n📊 ===== {dim_name} =====")
    result_map = {}
    for opt in options:
        # 每次先刷新清空筛选
        try:
            page.reload(wait_until="domcontentloaded")
            time.sleep(5)
        except Exception:
            time.sleep(3)
        # 展开维度浮层
        ok_open = click_label_open(page, dim_name)
        if not ok_open:
            log(f"  ⚠️ 无法展开「{dim_name}」浮层")
            # dump 当前可见文本帮助调试
            try:
                vis = page.evaluate("document.body.innerText.substring(0, 3000)")
                log(f"  页面文本(前3000字): {vis[:500]}")
            except:
                pass
            continue
        time.sleep(1.5)
        # 点击选项
        before = len(captured)
        ok = click_text(page, opt, wait=4.0)
        if not ok:
            # 尝试 XPath
            ok = click_xpath(page, f"//*[contains(text(),'{opt}')]", wait=4.0)
        if ok:
            fields = latest_filter_fields(before)
            if fields:
                # 看有没有指定 key
                if field_key and field_key in fields:
                    result_map[opt] = fields[field_key]
                    log(f"  ✅ {opt} → {field_key}={fields[field_key]}")
                else:
                    result_map[opt] = fields
                    log(f"  ✅ {opt} → {json.dumps(fields, ensure_ascii=False)[:120]}")
            else:
                log(f"  ⚠️ {opt} → 点了但没抓到新请求")
        else:
            log(f"  ❌ {opt} → 点击失败")
    RESULT[dim_name] = result_map
    log(f"  结果: {json.dumps(result_map, ensure_ascii=False)[:200]}")
    return result_map

def main():
    log("🔌 连接用户 Chrome (CDP 127.0.0.1:9222)...")
    pw = sync_playwright().start()
    browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
    ctx = browser.contexts[0]
    log(f"  上下文页面数: {len(ctx.pages)}")

    # 找达人广场页面
    page = None
    for p in ctx.pages:
        if "kwaixiaodian" in p.url or "daren" in p.url:
            page = p
            break
    if page is None:
        log("  未找到达人广场页面，新开标签页...")
        page = ctx.new_page()
        page.goto(DAREN_URL, timeout=60000, wait_until="domcontentloaded")
        time.sleep(12)
    log(f"  使用页面: {page.url[:100]}")
    log(f"  标题: {page.title()}")
    page.bring_to_front()
    time.sleep(3)

    # 先清空已有筛选
    log("🧹 清空已有筛选...")
    click_text(page, "清空", wait=2)
    try:
        page.reload(wait_until="domcontentloaded")
        time.sleep(5)
    except:
        pass

    # 拦截请求
    def on_request(request):
        if "promoter/list" in request.url and request.method == "POST":
            try:
                body = json.loads(request.post_data or "{}")
                captured.append(body)
            except Exception:
                pass
    page.on("request", on_request)

    # ============ 1. 粉丝年龄 ============
    capture_dimension(page, "粉丝年龄",
                      ["18岁以下", "18-23岁", "24-30岁", "31-40岁", "41-50岁", "50岁以上"],
                      field_key="fansAgeFeature")

    # ============ 2. 粉丝性别 ============
    capture_dimension(page, "粉丝性别",
                      ["女多男少", "男多女少", "男女均衡"],
                      field_key="fansGenderFeature")

    # ============ 3. 粉丝城市划分 ============
    capture_dimension(page, "粉丝城市划分",
                      ["一线城市", "新一线城市", "二线城市", "三线城市", "四线及以下"],
                      field_key="fansCityLevel")

    # ============ 4. 达人性别 ============
    capture_dimension(page, "达人性别",
                      ["男", "女"],
                      field_key="sellerGender")

    # ============ 5. 达人地域 ============
    capture_dimension(page, "达人地域",
                      ["北京", "上海", "广东"],
                      field_key="regionCode")

    # ============ 保存 ============
    dump = DEBUG_DIR / "cdp_filter_mapping.json"
    dump.write_text(json.dumps(RESULT, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"\n💾 映射结果: {dump}")

    raw = DEBUG_DIR / "cdp_all_bodies.json"
    raw.write_text(json.dumps(captured, ensure_ascii=False, indent=1), encoding="utf-8")
    log(f"💾 原始请求体: {raw} ({len(captured)} 个)")

    log("🏁 抓取完成。")
    try:
        browser.close()
    except:
        pass

if __name__ == "__main__":
    main()
