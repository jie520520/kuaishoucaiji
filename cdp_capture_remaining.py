# -*- coding: utf-8 -*-
"""补抓：粉丝性别（男粉丝多/女粉丝多）、达人性别、粉丝城市划分完整选项"""
import sys, json, time
from pathlib import Path
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except:
    pass
from playwright.sync_api import sync_playwright

DEBUG = Path(__file__).parent / "debug"
captured = []
RESULT = {}

def log(m): print(m, flush=True)

def on_req(req):
    if "promoter/list" in req.url and req.method == "POST":
        try:
            body = json.loads(req.post_data or "{}")
            captured.append(body)
        except:
            pass

pw = sync_playwright().start()
browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
ctx = browser.contexts[0]
page = None
for p in ctx.pages:
    if "kwaixiaodian" in p.url or "daren" in p.url:
        page = p
        break
if not page:
    page = ctx.new_page()
    page.goto("https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro", timeout=60000, wait_until="domcontentloaded")
    time.sleep(12)
page.bring_to_front()
page.on("request", on_req)

def do_filter(label, option, wait=4):
    """刷新→展开→点选项→抓请求"""
    page.reload(wait_until="domcontentloaded")
    time.sleep(5)
    # 点击维度标签
    try:
        page.locator(f"text={label}").first.click(timeout=8000)
        time.sleep(2)
    except:
        log(f"  无法展开 {label}")
        return None
    before = len(captured)
    try:
        page.locator(f"text={option}").first.click(timeout=8000)
        time.sleep(wait)
    except:
        log(f"  无法点击 {option}")
        return None
    if len(captured) > before:
        body = captured[-1]
        # 提取所有非常量字段
        special = {}
        for k, v in body.items():
            if k not in ("orderField", "orderType", "limit", "offset", "type", "hotSaleChannelIdList", "hotSaleSubChannelId"):
                special[k] = v
        return special
    return None

# ===== 粉丝性别 =====
log("\n=== 粉丝性别 ===")
fg = {}
for opt in ["男粉丝多", "女粉丝多"]:
    r = do_filter("粉丝性别", opt)
    if r:
        fg[opt] = r
        log(f"  ✅ {opt} → {json.dumps(r, ensure_ascii=False)}")
    else:
        log(f"  ❌ {opt} → 无结果")
RESULT["粉丝性别"] = fg

# ===== 达人性别 =====
log("\n=== 达人性别 ===")
# 先打开达人性别看选项
page.reload(wait_until="domcontentloaded")
time.sleep(5)
try:
    page.locator("text=达人性别").first.click(timeout=8000)
    time.sleep(2)
    # dump 选项
    opts = page.evaluate("""() => {
        const dd = document.querySelector('.rc-virtual-list');
        if (!dd) return [];
        return [...dd.querySelectorAll('*')].filter(el => {
            const t = el.textContent.trim();
            return t.length > 0 && t.length < 10 && el.getBoundingClientRect().width > 0;
        }).map(el => el.textContent.trim());
    }""")
    log(f"  达人性别选项: {opts}")
    # 逐个点击
    sg = {}
    for opt in opts:
        if not opt or opt in sg:
            continue
        page.reload(wait_until="domcontentloaded")
        time.sleep(5)
        try:
            page.locator("text=达人性别").first.click(timeout=8000)
            time.sleep(2)
        except:
            pass
        before = len(captured)
        try:
            page.locator(f"text={opt}").first.click(timeout=8000)
            time.sleep(4)
        except:
            pass
        if len(captured) > before:
            body = captured[-1]
            special = {}
            for k, v in body.items():
                if k not in ("orderField", "orderType", "limit", "offset", "type", "hotSaleChannelIdList", "hotSaleSubChannelId"):
                    special[k] = v
            sg[opt] = special
            log(f"  ✅ {opt} → {json.dumps(special, ensure_ascii=False)}")
        else:
            log(f"  ⚠️ {opt} → 无新请求")
    RESULT["达人性别"] = sg
except Exception as e:
    log(f"  达人性别探针失败: {e}")
    RESULT["达人性别"] = {}

# ===== 粉丝城市划分完整选项 =====
log("\n=== 粉丝城市划分（完整选项）===")
page.reload(wait_until="domcontentloaded")
time.sleep(5)
try:
    page.locator("text=粉丝城市划分").first.click(timeout=8000)
    time.sleep(2)
    city_opts = page.evaluate("""() => {
        const dd = document.querySelector('.rc-virtual-list');
        if (!dd) return [];
        return [...dd.querySelectorAll('*')].filter(el => {
            const t = el.textContent.trim();
            return t.length > 0 && t.length < 15 && el.getBoundingClientRect().width > 0;
        }).map(el => el.textContent.trim());
    }""")
    log(f"  城市划分选项: {city_opts}")
    cm = {}
    for opt in city_opts:
        if not opt or opt in cm:
            continue
        r = do_filter("粉丝城市划分", opt)
        if r:
            cm[opt] = r
            log(f"  ✅ {opt} → {json.dumps(r, ensure_ascii=False)}")
    RESULT["粉丝城市划分"] = cm
except Exception as e:
    log(f"  城市划分探针失败: {e}")

# ===== 有联系方式 =====
log("\n=== 有联系方式 ===")
page.reload(wait_until="domcontentloaded")
time.sleep(5)
before = len(captured)
try:
    page.locator("text=有联系方式").first.click(timeout=8000)
    time.sleep(4)
except:
    pass
if len(captured) > before:
    body = captured[-1]
    special = {}
    for k, v in body.items():
        if k not in ("orderField", "orderType", "limit", "offset", "type", "hotSaleChannelIdList", "hotSaleSubChannelId"):
            special[k] = v
    RESULT["有联系方式"] = special
    log(f"  ✅ {json.dumps(special, ensure_ascii=False)}")

# ===== 保存 =====
out = DEBUG / "cdp_filter_mapping_v2.json"
out.write_text(json.dumps(RESULT, ensure_ascii=False, indent=2), encoding="utf-8")
log(f"\n💾 补抓结果: {out}")
log(f"原始请求体: {len(captured)} 个")

try: browser.close()
except: pass
log("DONE")
