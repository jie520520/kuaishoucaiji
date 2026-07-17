# -*- coding: utf-8 -*-
"""
颜阿娇 - 批量提取达人联系方式 v4（修复版）
=================================
从采集结果Excel中读取达人列表，逐个打开详情页，
点击"查看联系方式"，提取手机号和微信号，输出到新Excel。

v4 修复内容（相对 v3）：
1. 【核心bug】微信号提取：无分隔符兜底正则原本要求首字符必须是字母，
   导致纯数字微信号（很常见，很多人直接用手机号当微信号）100%提取失败。
   现已放开，允许数字开头。
2. 【核心bug】手机号提取：原本用"排除4个以上连续相同数字"来过滤截断的系统ID，
   但这个方法会连带误杀真实存在的、恰好含有连续重复数字的手机号
   （如138****0000、159****1111这类很常见的号码）。
   现改用零宽断言 (?<!\\d)...(?!\\d) 精确匹配独立完整的11位号码，
   彻底避免误伤真实号码。
3. 【联动bug】旧版只要任一字段有值就跳过，导致历史漏提字段无法修复。
   现对没有可靠分类状态的历史部分结果复查一次；页面确认只有一种联系方式后永久跳过，
   若字段存在但解析失败，则保存诊断并当天不再查询。

技术路线（用户确认）：
- 点击"查看联系方式"后，手机号和微信号都可以直接文本复制
- 无需OCR，纯DOM文本提取即可

使用方式：
python extract_contacts.py --input 采集结果/颜阿娇_达人_健康_20260713_124840.xlsx
python extract_contacts.py --input 采集结果/颜阿娇_达人_健康_20260713_124840.xlsx --max 20
python extract_contacts.py --id 535256404 (单个测试)
"""
import sys
import io
import time
import json
import re
import argparse
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        if hasattr(sys.stderr, 'buffer') and not isinstance(sys.stderr, io.TextIOWrapper):
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(line_buffering=True)
    except Exception:
        pass  # 作为模块导入或管道模式时忽略

_START_TIME = time.time()


def _ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


def log(msg: str, end: str = "\n"):
    text = f"[{_ts()}] {msg}"
    print(text, end=end, flush=True)


from playwright.sync_api import sync_playwright
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# ============================================================
# 配置
# ============================================================
class Config:
    BASE_DIR = Path(__file__).parent
    USER_DATA = BASE_DIR / ".kuaishou_browser_data"
    OUTPUT_DIR = BASE_DIR / "采集结果"
    DEBUG_DIR = BASE_DIR / "debug_screenshots"

    DAREN_SQUARE_URL = "https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro"
    DETAIL_URL_FMT = "https://cps.kwaixiaodian.com/zone/daren-match/daren-detail?promoterId={pid}"

    # 提取限制
    MIN_PAGE_DELAY = 10
    PAGE_DELAY = 10  # 每人间隔（秒）
    MAX_PER_SESSION = 500  # 每次最多提取人数（官方每日上限 500）
    BATCH_SIZE = 10  # 每批人数（中间暂停一下）

    # 并发/频率限制：保守策略
    LONG_PAUSE_EVERY = 10  # 每N人后额外暂停
    LONG_PAUSE_DURATION = 30  # 额外暂停时长（秒）

    # 自动暂停恢复：连续N个达人无联系方式 → 判断触达每日上限 → 等30分钟恢复
    AUTO_PAUSE_CONSECUTIVE_EMPTY = 5  # 连续N个达人无联系方式触发暂停（5=较合理，避免偶发空结果误触发）
    AUTO_PAUSE_MINUTES = 30  # 暂停分钟后自动恢复
    MAX_CONTACT_RETRIES = 1


_MISSING_CONTACT_VALUES = {"", "none", "无", "暂无", "未获取"}

def _clean_contact_value(value) -> str:
    text = str(value).strip() if value is not None else ""
    return "" if text.lower() in _MISSING_CONTACT_VALUES else text

def _is_placeholder_contact(value) -> bool:
    text = str(value).strip() if value is not None else ""
    return bool(text) and text.lower() in _MISSING_CONTACT_VALUES

def _contact_values_from_row(row, phone_col, wechat_col):
    phone = _clean_contact_value(row[phone_col]) if phone_col is not None and len(row) > phone_col else ""
    wechat = _clean_contact_value(row[wechat_col]) if wechat_col is not None and len(row) > wechat_col else ""
    return phone, wechat

def _clear_placeholder_contacts(input_excel: str):
    wb = load_workbook(input_excel)
    ws = wb.active
    contact_cols = []
    for col_idx, cell in enumerate(ws[1], 1):
        header = str(cell.value).strip() if cell.value else ""
        if header in ("手机号", "微信号"):
            contact_cols.append(col_idx)

    changed = 0
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in contact_cols:
            cell = ws.cell(row=row_idx, column=col_idx)
            if _is_placeholder_contact(cell.value):
                cell.value = None
                changed += 1

    if changed:
        wb.save(input_excel)
        log(f"   🧹 已清理 {changed} 个旧的‘无/暂无’占位值")
    wb.close()
    return changed

def _retry_state_path(input_excel: str) -> Path:
    input_path = Path(input_excel)
    return input_path.with_name(f"{input_path.stem}_联系方式重试.json")

def _load_retry_state(input_excel: str) -> dict:
    state_path = _retry_state_path(input_excel)
    if not state_path.exists():
        return {"input": str(Path(input_excel).resolve()), "creators": {}}
    try:
        with open(state_path, "r", encoding="utf-8") as file:
            state = json.load(file)
        if not isinstance(state, dict) or not isinstance(state.get("creators"), dict):
            raise ValueError("重试状态格式无效")
        return state
    except Exception as error:
        log(f"⚠️ 重试状态读取失败，将重新建立: {error}")
        return {"input": str(Path(input_excel).resolve()), "creators": {}}

def _save_retry_state(input_excel: str, state: dict):
    state_path = _retry_state_path(input_excel)
    state["input"] = str(Path(input_excel).resolve())
    state["max_retries"] = Config.MAX_CONTACT_RETRIES
    state["updated_at"] = datetime.now().isoformat()
    temp_path = state_path.with_suffix(state_path.suffix + ".tmp")
    with open(temp_path, "w", encoding="utf-8") as file:
        json.dump(state, file, ensure_ascii=False, indent=2)
    temp_path.replace(state_path)

_TERMINAL_CONTACT_STATUSES = {"完整", "仅手机号", "仅微信号"}
_RETRYABLE_CONTACT_STATUSES = {"解析异常", "平台临时未显示", "技术失败", "采集标签与详情冲突", "未填写联系方式", "今日暂缓", "待重试"}

def _is_retry_exhausted(retry_state: dict, promoter_id: str) -> bool:
    entry = retry_state.get("creators", {}).get(str(promoter_id), {})
    today = datetime.now().date().isoformat()
    return (
        entry.get("last_attempt_date") == today
        and entry.get("status") in _RETRYABLE_CONTACT_STATUSES
        and int(entry.get("failures", 0)) >= Config.MAX_CONTACT_RETRIES
    )

def _is_contact_complete(retry_state: dict, promoter_id: str, phone: str, wechat: str) -> bool:
    entry = retry_state.get("creators", {}).get(str(promoter_id), {})
    status = entry.get("status")
    if status in _TERMINAL_CONTACT_STATUSES:
        return True
    if status in _RETRYABLE_CONTACT_STATUSES:
        return False
    if phone and wechat:
        return True
    return False

def _record_contact_attempt(input_excel: str, retry_state: dict, promoter_id: str,
                            existing_phone: str, existing_wechat: str, result: dict) -> dict:
    pid = str(promoter_id).strip()
    new_phone = _clean_contact_value(result.get("phone"))
    new_wechat = _clean_contact_value(result.get("wechat"))
    final_phone = existing_phone or new_phone
    final_wechat = existing_wechat or new_wechat
    today = datetime.now().date().isoformat()
    outcome = result.get("query_outcome", "technical_failure")

    status_by_outcome = {
        "complete": "完整",
        "only_phone": "仅手机号",
        "only_wechat": "仅微信号",
        "source_conflict": "采集标签与详情冲突",
        "parse_failure": "解析异常",
        "temporary_failure": "平台临时未显示",
        "technical_failure": "技术失败",
    }
    status = status_by_outcome.get(outcome, "技术失败")
    failures = 0 if status in _TERMINAL_CONTACT_STATUSES else Config.MAX_CONTACT_RETRIES

    creators = retry_state.setdefault("creators", {})
    creators[pid] = {
        "failures": failures,
        "status": status,
        "query_outcome": outcome,
        "phone": final_phone,
        "wechat": final_wechat,
        "last_error": result.get("error", ""),
        "classification_reason": result.get("classification_reason", ""),
        "diagnostic_text": result.get("diagnostic_text", ""),
        "screenshot_path": result.get("screenshot_path", ""),
        "response_wait_timed_out": result.get("response_wait_timed_out", False),
        "rate_limit_warning": result.get("rate_limit_warning", ""),
        "last_attempt_date": today,
        "updated_at": datetime.now().isoformat(),
    }
    _save_retry_state(input_excel, retry_state)

    result["phone"] = new_phone
    result["wechat"] = new_wechat
    result["retry_failures"] = failures
    result["retry_status"] = status
    result["final_phone"] = final_phone
    result["final_wechat"] = final_wechat
    return result


# ============================================================
# 单个达人联系方式提取
# ============================================================
def _contact_text_excerpt(text: str, limit: int = 800) -> str:
    normalized = re.sub(r'\s+', ' ', text or '').strip()
    if len(normalized) <= limit:
        return normalized
    positions = [normalized.find(marker) for marker in ("联系方式", "手机号", "微信号", "操作频繁", "稍后再试")]
    positions = [position for position in positions if position >= 0]
    start = max(0, (min(positions) if positions else 0) - 120)
    return normalized[start:start + limit]

def _read_contact_text(page, body_text: str) -> str:
    for selector in (
        '[role="dialog"]',
        '.ant-modal-content',
        '.semi-modal-content',
        '.el-dialog',
        '[class*="contact"][class*="modal"]',
    ):
        try:
            locator = page.locator(selector)
            if locator.count() > 0:
                panel_text = locator.last.inner_text(timeout=2000).strip()
                if panel_text:
                    return panel_text
        except Exception:
            continue
    return body_text

def _find_contact_button(page):
    candidates = []
    try:
        candidates.append(page.get_by_role("button", name="查看联系方式", exact=True))
    except Exception:
        pass
    try:
        candidates.append(page.get_by_text("查看联系方式", exact=True))
    except Exception:
        pass
    candidates.append(page.locator('text=查看联系方式'))
    for locator in candidates:
        try:
            if locator.count() > 0:
                return locator.first
        except Exception:
            continue
    return None

def _wait_for_contact_response(page, before_text: str) -> bool:
    try:
        page.wait_for_function(
            """before => {
                const text = document.body.innerText || '';
                const markers = ['手机号', '手机号码', '微信号', '操作频繁', '稍后再试',
                    '系统繁忙', '操作过快', '请稍后重试', '暂无联系方式', '未填写联系方式', '安全检测', '滑动验证'];
                return text !== before && markers.some(marker => text.includes(marker));
            }""",
            arg=before_text,
            timeout=10000,
        )
        return True
    except Exception:
        return False

def _save_failure_screenshot(page, promoter_id: str, outcome: str) -> str:
    try:
        Config.DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_pid = re.sub(r'[^a-zA-Z0-9_-]', '_', str(promoter_id))
        shot_path = Config.DEBUG_DIR / f"{safe_pid}_{outcome}_{timestamp}.png"
        page.screenshot(path=str(shot_path), full_page=True)
        return str(shot_path)
    except Exception:
        return ""

def _detect_verification(page) -> str:
    try:
        text = page.locator("body").inner_text(timeout=2000)
    except Exception:
        return ""
    for marker in ("滑动验证", "请完成验证", "安全检测", "人机验证", "验证码", "拖动滑块"):
        if marker in text:
            return marker
    return ""

def extract_one(page, promoter_id: str, nickname: str = "") -> dict:
    result = {
        "promoterId": promoter_id,
        "nickname": nickname,
        "phone": "",
        "wechat": "",
        "error": "",
        "query_outcome": "technical_failure",
        "classification_reason": "",
        "diagnostic_text": "",
        "screenshot_path": "",
        "response_wait_timed_out": False,
        "rate_limit_warning": "",
    }
    detail_url = Config.DETAIL_URL_FMT.format(pid=promoter_id)

    try:
        page.goto(detail_url, timeout=30000, wait_until="domcontentloaded")
        time.sleep(4)

        body_text = page.locator("body").inner_text(timeout=3000)
        if "未找到" in body_text or "404" in body_text:
            result["error"] = "详情页不存在"
            result["classification_reason"] = "详情页不存在，未消耗为有效联系方式结果"
            return result

        contact_btn = _find_contact_button(page)
        if contact_btn is not None:
            try:
                before_click_text = body_text
                contact_btn.click(timeout=5000)
                response_loaded = _wait_for_contact_response(page, before_click_text)
                result["response_wait_timed_out"] = not response_loaded
            except Exception as error:
                result["error"] = f"点击查看联系方式失败: {error}"
                result["classification_reason"] = "按钮点击失败，不能判断联系方式状态"
                result["screenshot_path"] = _save_failure_screenshot(page, promoter_id, "click_failure")
                return result
        elif not any(marker in body_text for marker in ("手机号", "手机号码", "微信号", "暂无联系方式", "未填写联系方式")):
            result["error"] = "无查看联系方式按钮"
            result["classification_reason"] = "页面未出现按钮或已展开联系方式区域"
            result["screenshot_path"] = _save_failure_screenshot(page, promoter_id, "button_missing")
            return result

        body_text = page.locator("body").inner_text(timeout=3000)
        contact_text = _read_contact_text(page, body_text)
        result["diagnostic_text"] = _contact_text_excerpt(contact_text)

        phone_patterns = [
            r'手机号\s*[：:]?\s*(?:\+?86\s*)?(1[3-9]\d{9})(?!\d)',
            r'手机号码\s*[：:]?\s*(?:\+?86\s*)?(1[3-9]\d{9})(?!\d)',
            r'联系电话\s*[：:]?\s*(?:\+?86\s*)?(1[3-9]\d{9})(?!\d)',
        ]
        phones = []
        for pattern in phone_patterns:
            phones.extend(re.findall(pattern, contact_text))
        phones = list(dict.fromkeys(phones))
        result["phone"] = phones[0] if phones else ""

        wechat_patterns = [
            r'微信号\s*[：:]\s*([a-zA-Z0-9_-]{4,40})',
            r'微信\s*[：:]\s*([a-zA-Z0-9_-]{4,40})',
            r'微信号\s+([a-zA-Z0-9_-]{4,40})',
            r'微信\s+([a-zA-Z0-9_-]{4,40})',
            r'微信号([a-zA-Z0-9][a-zA-Z0-9_-]{3,39})',
            r'微信([a-zA-Z0-9][a-zA-Z0-9_-]{3,39})',
        ]
        wechats = []
        for pattern in wechat_patterns:
            wechats.extend(re.findall(pattern, contact_text))
        excluded_wechat = {"该用户", "暂无", "null", "undefined", "wxid", "gh_", "未填写", "未设置"}
        wechats = [
            value for value in dict.fromkeys(wechats)
            if value.lower() not in excluded_wechat and not value.startswith("****")
        ]
        result["wechat"] = wechats[0] if wechats else ""

        phone_label_present = bool(re.search(r'手机号|手机号码|联系电话', contact_text))
        wechat_label_present = bool(re.search(r'微信号|微信号码|微信', contact_text))
        empty_value_pattern = r'\s*[：:]?\s*(?:--+|—+|无|暂无|未填写|未设置|未提供)'
        phone_explicitly_empty = bool(re.search(
            rf'(?:手机号|手机号码|联系电话){empty_value_pattern}', contact_text
        ))
        wechat_explicitly_empty = bool(re.search(
            rf'(?:微信号|微信号码|微信){empty_value_pattern}', contact_text
        ))
        temporary_markers = (
            "操作过快", "操作频繁", "请求频繁", "访问频繁", "请稍后重试",
            "稍后再试", "系统繁忙", "网络异常", "查询次数",
            "次数已达上限", "今日上限", "加载失败",
        )
        result["rate_limit_warning"] = next(
            (marker for marker in temporary_markers if marker in body_text), ""
        )
        explicit_empty_markers = (
            "未填写联系方式", "暂无联系方式", "未设置联系方式",
            "没有填写联系方式", "未提供联系方式",
        )

        if result["phone"] and result["wechat"]:
            result["query_outcome"] = "complete"
            result["classification_reason"] = "手机号和微信号均已明确解析"
        elif result["phone"]:
            if wechat_explicitly_empty or not wechat_label_present:
                result["query_outcome"] = "only_phone"
                result["classification_reason"] = "页面明确未填写微信号，或未展示微信字段"
            else:
                result["query_outcome"] = "parse_failure"
                result["error"] = "页面出现微信字段，但未解析到微信号"
                result["classification_reason"] = "保留手机号；微信字段可能格式变化，需根据诊断文本修复"
        elif result["wechat"]:
            if phone_explicitly_empty or not phone_label_present:
                result["query_outcome"] = "only_wechat"
                result["classification_reason"] = "页面明确未填写手机号，或未展示手机字段"
            else:
                result["query_outcome"] = "parse_failure"
                result["error"] = "页面出现手机字段，但未解析到手机号"
                result["classification_reason"] = "保留微信号；手机字段可能格式变化，需根据诊断文本修复"
        elif (phone_explicitly_empty and wechat_explicitly_empty) or any(marker in contact_text for marker in explicit_empty_markers):
            result["query_outcome"] = "source_conflict"
            result["error"] = "采集阶段标记有联系方式，但详情页未显示"
            result["classification_reason"] = "采集标签与详情结果冲突，不能判定达人未填写"
        elif any(marker in body_text for marker in temporary_markers):
            result["query_outcome"] = "temporary_failure"
            result["error"] = "平台临时未显示联系方式"
            result["classification_reason"] = "页面出现频繁、上限、繁忙或加载失败提示"
        elif phone_label_present or wechat_label_present:
            result["query_outcome"] = "parse_failure"
            result["error"] = "页面出现联系方式字段，但未解析到有效值"
            result["classification_reason"] = "疑似页面格式变化，已保存诊断文本"
        else:
            result["query_outcome"] = "temporary_failure"
            result["error"] = "联系方式区域未完整显示"
            result["classification_reason"] = "没有明确未填写提示，不能判定达人无联系方式"

    except Exception as error:
        result["error"] = str(error)[:200]
        result["query_outcome"] = "technical_failure"
        result["classification_reason"] = "页面导航、读取或解析发生技术异常"

    if result["query_outcome"] not in ("complete", "only_phone", "only_wechat"):
        result["screenshot_path"] = _save_failure_screenshot(page, promoter_id, result["query_outcome"])
    return result


# ============================================================
# 浏览器生命周期 & 批量提取
# ============================================================
def _open_browser():
    """打开浏览器并返回 (p, ctx, page)，同时导航空到达人广场确认登录。"""
    log("🚀 启动浏览器（使用已有登录态）...")
    p = sync_playwright().start()
    ctx = p.chromium.launch_persistent_context(
        user_data_dir=str(Config.USER_DATA),
        headless=False,
        args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        viewport={"width": 1440, "height": 900},
        locale="zh-CN",
    )
    page = ctx.new_page()
    page.goto(Config.DAREN_SQUARE_URL, timeout=60000, wait_until="domcontentloaded")
    time.sleep(5)

    if "login" in page.url.lower():
        log("⚠️ 需要扫码登录快手，请在浏览器中扫码...")
        for i in range(120):
            time.sleep(2)
            if "login" not in page.url.lower():
                log("🔓 登录成功！")
                time.sleep(3)
                break
        else:
            log("❌ 登录超时")
            ctx.close()
            p.stop()
            return None, None, None

    return p, ctx, page


def _close_browser(p, ctx):
    """安全关闭浏览器"""
    try:
        ctx.close()
    except Exception:
        pass
    try:
        p.stop()
    except Exception:
        pass


def batch_extract(input_excel: str, max_count: int = None,
                   start_from: int = 0, callback=None):
    """
    批量从Excel中读取达人列表，逐个提取联系方式。

    特性：
    - 智能跳过已完成分类的达人（断点续提）
    - 历史部分结果仅复查一次，用于补回旧版正则漏掉的数据
    - 页面确认仅有一种联系方式后不再查询，解析异常则保存诊断并当天跳过
    - 连续N个达人无联系方式 → 自动暂停30分钟 → 恢复（模拟人工等待触达上限恢复）

    参数：
        input_excel: 采集结果Excel路径
        max_count: 最多提取人数（None=不限制，默认500上限）
        start_from: 从第几个待提取达人开始（用于断点续传）
        callback: 回调函数 callback(row_index, result, progress_info)

    返回：
        list of dict: 所有提取结果
    """
    log(f"📂 读取输入文件: {input_excel}")
    wb = load_workbook(input_excel, read_only=True)
    ws = wb.active

    # 读取表头，找到关键列
    header = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    header = [str(h).strip() if h else "" for h in header]

    pid_col = None
    name_col = None
    phone_col = None
    wechat_col = None
    for i, h in enumerate(header):
        hl = h.lower()
        if "快手id" in hl or "promoter" in hl or h == "快手ID":
            pid_col = i
        if "达人名称" in hl or "昵称" in hl or h == "达人名称":
            name_col = i
        if "手机号" in h:
            phone_col = i
        if "微信号" in h:
            wechat_col = i

    if pid_col is None:
        log("❌ 未找到'快手ID'或'promoterId'列")
        log(f"   表头: {header}")
        wb.close()
        return []

    log(f"   promoterId列: {pid_col} ({header[pid_col]})")
    if name_col is not None:
        log(f"   昵称列: {name_col} ({header[name_col]})")

    # 读取全部数据行，标记哪些已有完整联系方式
    all_rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    _clear_placeholder_contacts(input_excel)

    if not all_rows:
        log("❌ Excel中没有数据行")
        return []

    retry_state = _load_retry_state(input_excel)
    rows_to_extract = []
    already_has = 0
    retry_exhausted = 0
    for idx, row in enumerate(all_rows):
        promoter_id = str(row[pid_col]).strip() if row[pid_col] else ""
        phone_val, wechat_val = _contact_values_from_row(row, phone_col, wechat_col)

        if promoter_id and _is_retry_exhausted(retry_state, promoter_id):
            retry_exhausted += 1
        elif _is_contact_complete(retry_state, promoter_id, phone_val, wechat_val):
            already_has += 1
        else:
            rows_to_extract.append((idx + 2, row, phone_val, wechat_val))

    total = len(all_rows)
    need_extract = len(rows_to_extract)
    if already_has > 0:
        log(f"   ⏭ 已完成或已可靠分类: {already_has} 人 (自动跳过)")
    if retry_exhausted > 0:
        log(f"   🛡 今日已查询但结果不可靠: {retry_exhausted} 人 (当天不再重复消耗名额)")
    log(f"   🎯 待提取: {need_extract} 人")

    if need_extract == 0:
        log("✅ 所有达人已完成分类，或今日已达到安全查询上限")
        return []

    if max_count is None:
        max_count = min(Config.MAX_PER_SESSION, need_extract)
    else:
        max_count = min(max_count, need_extract)

    if start_from >= need_extract:
        log(f"❌ start_from={start_from} 超出待提取范围")
        return []

    effective_total = min(max_count, need_extract - start_from)
    log(f"📊 本次提取: {start_from+1} ~ {start_from+effective_total} / {need_extract} (共{effective_total}人)")

    # 启动浏览器
    p, ctx, page = _open_browser()
    if page is None:
        return []

    results = []
    success_count = 0
    consecutive_empty = 0  # 连续无联系方式的计数
    pause_count = 0  # 暂停次数

    # 使用 while 循环支持暂停恢复
    idx = start_from
    while idx < start_from + effective_total:
        row_excel_num, row, existing_phone, existing_wechat = rows_to_extract[idx]
        promoter_id = str(row[pid_col]) if row[pid_col] else ""
        nickname = str(row[name_col]) if name_col is not None and len(row) > name_col and row[name_col] else ""

        if not promoter_id or promoter_id.upper() == "NONE":
            log(f"   ⏭ [{idx+1}/{need_extract}] 跳过：无ID")
            results.append({"promoterId": "", "nickname": nickname, "phone": "", "wechat": "", "error": "无ID"})
            idx += 1
            continue

        # 进度显示
        done_so_far = idx - start_from + 1
        remaining = effective_total - done_so_far
        progress = f"[{idx+1}/{need_extract}] 已提取{success_count}人, 剩余{remaining}人"
        log(f"\n{'='*50}")
        log(f"   🎯 {progress}")
        log(f"   📋 {nickname} (ID:{promoter_id})")

        # 提取
        result = extract_one(page, promoter_id, nickname)
        result = _record_contact_attempt(
            input_excel, retry_state, promoter_id, existing_phone, existing_wechat, result
        )

        if result["phone"] or result["wechat"]:
            success_count += 1
            consecutive_empty = 0
            emoji = "✅"
        else:
            consecutive_empty += 1
            emoji = "⚠️ "

        log(f"   {emoji} 手机: {result['phone'] or '---'}  微信: {result['wechat'] or '---'}")
        if result["error"]:
            log(f"   错误: {result['error']}")
        if result["retry_status"] == "解析异常":
            log("   🧩 页面有字段但未解析成功；已保存诊断文本，当天不再查询")
        elif result["retry_status"] == "平台临时未显示":
            log("   🛡 平台临时未显示；当天不再查询，明日可重试")
        elif result["retry_status"] == "技术失败":
            log("   🛡 技术异常；当天不再自动查询，避免重复消耗名额")
        elif result["retry_status"] == "采集标签与详情冲突":
            log("   🛡 采集标签确认有联系方式，但详情页未显示；当天不再查询")

        results.append(result)

        # 回调
        if callback:
            callback(idx, result, progress)

        # 保存进度：JSON 每步存，Excel 每5步批量写（防卡锁）
        _save_progress(results, input_excel)
        if done_so_far % 5 == 0 or done_so_far == effective_total or result["phone"] or result["wechat"]:
            try:
                _save_final_excel(results, input_excel)
            except Exception as e:
                log(f"   ⚠️ Excel保存失败（JSON已存）: {e}")

        idx += 1  # 指针前移（无论失败与否，不重试同一个达人）

        # ========================================
        # 检测：连续N个无联系方式 → 自动暂停恢复
        # ========================================
        if result.get("rate_limit_warning") or consecutive_empty >= Config.AUTO_PAUSE_CONSECUTIVE_EMPTY:
            pause_count += 1
            log(f"\n{'⚠️'*20}")
            if result.get("rate_limit_warning"):
                log(f"⚠️ 检测到平台提示：{result['rate_limit_warning']}")
                log("⚠️ 当前达人结果已保存，立即暂停，不再打开下一位达人")
            else:
                log(f"⚠️ 连续 {consecutive_empty} 个达人未可靠显示联系方式！")
                log("⚠️ 疑似触达查询限制，触发自动暂停...")
            log(f"{'⚠️'*20}")

            # 先保存已有结果到 Excel
            _save_final_excel(results, input_excel)

            verification_marker = _detect_verification(page)
            if verification_marker:
                log(f"🧩 检测到‘{verification_marker}’，请在浏览器中人工完成验证")
                log("   验证期间不会继续打开达人页面，避免浪费查看名额")
                resolved = False
                for _ in range(150):
                    time.sleep(2)
                    if not _detect_verification(page):
                        resolved = True
                        break
                if not resolved:
                    log("❌ 5分钟内未完成验证，本次任务安全停止")
                    break
                log("✅ 验证已完成，继续提取")
                consecutive_empty = 0
                continue

            # 不关浏览器！保持登录会话存活，避免恢复时需要人工扫码
            page.goto(Config.DAREN_SQUARE_URL, timeout=30000, wait_until="domcontentloaded")
            time.sleep(2)

            wait_seconds = Config.AUTO_PAUSE_MINUTES * 60
            wait_end = datetime.now().timestamp() + wait_seconds
            log(f"⏰ 暂停 {Config.AUTO_PAUSE_MINUTES} 分钟（浏览器保持打开，无需重新登录）")
            log(f"   预计 {datetime.fromtimestamp(wait_end).strftime('%H:%M:%S')} 恢复...")
            log(f"   (已提取 {len(results)} 人，成功 {success_count} 人，第 {pause_count} 次暂停)")

            # 分段等待，每30秒输出一次剩余时间
            while wait_seconds > 0:
                chunk = min(30, wait_seconds)
                time.sleep(chunk)
                wait_seconds -= chunk
                if wait_seconds > 0:
                    log(f"   ⏳ 剩余约 {wait_seconds // 60} 分 {wait_seconds % 60} 秒...")

            # 快速验证登录态（导航到广场确认）
            log(f"\n{'='*50}")
            log(f"🔄 第 {pause_count} 次恢复：验证登录态...")
            page.goto(Config.DAREN_SQUARE_URL, timeout=30000, wait_until="domcontentloaded")
            time.sleep(3)

            if "login" in page.url.lower():
                log("⚠️ 登录态已过期，需要重新扫码（等待 30 秒自动检测）...")
                for _ in range(15):
                    time.sleep(2)
                    if "login" not in page.url.lower():
                        log("🔓 登录已恢复！")
                        break
                else:
                    log("❌ 登录态已过期且无法自动恢复，请手动扫码后继续")
                    log("   浏览器窗口已打开，扫码后程序将继续...")
                    for _ in range(60):
                        time.sleep(2)
                        if "login" not in page.url.lower():
                            log("🔓 登录成功，继续提取！")
                            break
                    else:
                        log("❌ 登录超时，停止提取")
                        break

            consecutive_empty = 0  # 重置计数器
            log(f"✅ 已恢复，从第 {idx-start_from+1}/{effective_total} 个达人继续")
            continue  # 不需要 sleep，直接下一个

        # 频率控制（正常模式）
        if idx < start_from + effective_total:
            delay = Config.PAGE_DELAY
            log(f"   ⏳ 等待 {delay} 秒...")
            time.sleep(delay)

            # 每N人额外暂停
            if done_so_far % Config.LONG_PAUSE_EVERY == 0:
                log(f"   🫁 每{Config.LONG_PAUSE_EVERY}人额外暂停 {Config.LONG_PAUSE_DURATION} 秒...")
                time.sleep(Config.LONG_PAUSE_DURATION)

    # 关闭浏览器
    _close_browser(p, ctx)

    # 保存最终结果
    output_file = _save_final_excel(results, input_excel)

    log(f"\n{'='*50}")
    log(f"🎉 提取完成！")
    log(f"   成功: {success_count}/{len(results)}")
    if pause_count > 0:
        log(f"   自动暂停恢复: {pause_count} 次")
    log(f"   输出: {output_file}")

    return results


# ============================================================
# 进度保存 & 最终输出
# ============================================================
def _save_progress(results: list, input_excel: str):
    """增量保存为JSON（用于断点续传恢复）"""
    input_stem = Path(input_excel).stem
    progress_file = Config.OUTPUT_DIR / f"{input_stem}_联系方式_progress.json"
    with open(progress_file, "w", encoding="utf-8") as f:
        json.dump({
            "input": str(input_excel),
            "extracted": len(results),
            "timestamp": datetime.now().isoformat(),
            "results": results,
        }, f, ensure_ascii=False, indent=2)


def _save_final_excel(results: list, input_excel: str):
    """将提取的手机号和微信号回写到原始Excel。

    【v4修复】已有的有效字段保留不覆盖；缺失的字段（哪怕只缺一个）
    用本次提取结果补上。之前版本是"任一字段有值就整行跳过"，
    会导致"手机号有、微信号空"的行永远补不上微信号。
    """
    input_path = Path(input_excel)

    # 构建结果查找表：{promoterId: (phone, wechat)}
    # 有联系方式的：写实际号码；无联系方式的：phone写"无"作标记，避免下次重复提取
    lookup = {}
    for r in results:
        pid = str(r.get("promoterId", "")).strip()
        if not pid:
            continue
        phone = _clean_contact_value(r.get("final_phone") or r.get("phone"))
        wechat = _clean_contact_value(r.get("final_wechat") or r.get("wechat"))
        lookup[pid] = (phone, wechat)

    if not lookup:
        log("⚠️ 没有可回写的数据")
        return ""

    # 打开原始文件（读写模式）
    wb = load_workbook(input_excel)
    ws = wb.active

    # 读取表头，找关键列
    pid_col = None
    phone_col = None
    wechat_col = None
    for col_idx, cell in enumerate(ws[1], 1):
        h = str(cell.value).strip() if cell.value else ""
        if "快手ID" in h or "promoter" in h.lower():
            pid_col = col_idx
        if h == "手机号":
            phone_col = col_idx
        if h == "微信号":
            wechat_col = col_idx

    if pid_col is None:
        log("❌ 未找到快手ID列，无法回写到原文件")
        wb.close()
        return ""

    # 如果列不存在则新增，存在则复用已有列号
    need_new_phone = (phone_col is None)
    need_new_wechat = (wechat_col is None)
    max_col = ws.max_column
    if need_new_phone and need_new_wechat:
        phone_col = max_col + 1
        wechat_col = max_col + 2
    elif need_new_phone:
        phone_col = max_col + 1
    elif need_new_wechat:
        wechat_col = max_col + 1

    # 样式
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="微软雅黑", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # 写新表头（仅新建列需要）
    for col, title, is_new in [(phone_col, "手机号", need_new_phone), (wechat_col, "微信号", need_new_wechat)]:
        if is_new:
            cell = ws.cell(row=1, column=col, value=title)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

    # 逐行回填：已有的有效字段保留，缺失的字段才用新结果补上
    filled = 0
    skipped = 0
    max_row = ws.max_row
    for row_idx in range(2, max_row + 1):
        pid_cell = ws.cell(row=row_idx, column=pid_col)
        pid = str(pid_cell.value).strip() if pid_cell.value else ""

        phone_cell = ws.cell(row=row_idx, column=phone_col)
        wechat_cell = ws.cell(row=row_idx, column=wechat_col)
        existing_phone = phone_cell.value
        existing_wechat = wechat_cell.value
        if _is_placeholder_contact(existing_phone):
            phone_cell.value = None
        if _is_placeholder_contact(existing_wechat):
            wechat_cell.value = None
        existing_phone_str = _clean_contact_value(existing_phone)
        existing_wechat_str = _clean_contact_value(existing_wechat)

        has_good_phone = bool(existing_phone_str)
        has_good_wechat = bool(existing_wechat_str)

        if has_good_phone and has_good_wechat:
            # 两个字段都已经是有效数据，整行跳过，不覆盖
            skipped += 1
            continue

        # 从本次提取结果查找
        phone_val, wechat_val = lookup.get(pid, ("", ""))
        if not phone_val and not wechat_val:
            # 这个pid本次没有新结果，保留原样
            continue

        final_phone = existing_phone_str if has_good_phone else phone_val
        final_wechat = existing_wechat_str if has_good_wechat else wechat_val
        has_real_contact = bool((final_phone and final_phone != "无") or final_wechat)

        if has_real_contact:
            filled += 1

        for col, val in [(phone_col, final_phone), (wechat_col, final_wechat)]:
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.font = body_font
            cell.border = thin_border
            cell.alignment = center_align

        # 颜色：绿=有联系方式，红=无/未提取到
        row_fill = green_fill if has_real_contact else red_fill
        ws.cell(row=row_idx, column=phone_col).fill = row_fill
        ws.cell(row=row_idx, column=wechat_col).fill = row_fill

    # 列宽
    from openpyxl.utils import get_column_letter
    ws.column_dimensions[get_column_letter(phone_col)].width = 16
    ws.column_dimensions[get_column_letter(wechat_col)].width = 22

    # 保存回原文件
    wb.save(input_excel)
    wb.close()

    log(f"💾 已回写到原文件: {input_path.name}")
    log(f"   本次新增/补全: {filled} 人")
    if skipped > 0:
        log(f"   已有完整数据跳过: {skipped} 人 (保留原数据)")

    return str(input_path)


# ============================================================
# 单个结果精准追加（用于实时保存，O(1)行写入）
# ============================================================
# 缓存：{input_excel: (pid_col, phone_col, wechat_col)}，避免每次检测列
_col_cache = {}


def _save_one_to_excel(result: dict, input_excel: str):
    """将单个提取结果精准写入Excel对应行。

    与 _save_final_excel 的区别：
    - _save_final_excel 遍历全部行再写 → O(n^2) I/O，日志混淆
    - _save_one_to_excel 按PID精准定位→只写一行 → O(1)，日志清晰

    【v4修复】同样改为"已有的有效字段保留，缺失的字段才补"，
    而不是"任一字段有值就整行跳过"。
    """
    pid = str(result.get("promoterId", "")).strip()
    if not pid:
        return ""

    phone = _clean_contact_value(result.get("final_phone") or result.get("phone"))
    wechat = _clean_contact_value(result.get("final_wechat") or result.get("wechat"))
    has_real = bool(phone or wechat)

    input_path = Path(input_excel)
    key = str(input_path.resolve())

    wb = load_workbook(input_excel)
    ws = wb.active

    # 列检测（优先用缓存）
    cached = _col_cache.get(key)
    if cached:
        pid_col, phone_col, wechat_col = cached
    else:
        pid_col = phone_col = wechat_col = None
        for col_idx, cell in enumerate(ws[1], 1):
            h = str(cell.value).strip() if cell.value else ""
            if "快手ID" in h or "promoter" in h.lower():
                pid_col = col_idx
            if h == "手机号":
                phone_col = col_idx
            if h == "微信号":
                wechat_col = col_idx

        if pid_col is None:
            wb.close()
            return ""

        need_new_phone = (phone_col is None)
        need_new_wechat = (wechat_col is None)
        max_col = ws.max_column
        if need_new_phone and need_new_wechat:
            phone_col = max_col + 1
            wechat_col = max_col + 2
        elif need_new_phone:
            phone_col = max_col + 1
        elif need_new_wechat:
            wechat_col = max_col + 1

        # 写新表头
        header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                              top=Side(style="thin"), bottom=Side(style="thin"))
        for col, title, is_new in [(phone_col, "手机号", need_new_phone), (wechat_col, "微信号", need_new_wechat)]:
            if is_new:
                cell = ws.cell(row=1, column=col, value=title)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

        _col_cache[key] = (pid_col, phone_col, wechat_col)

    # 样式
    body_font = Font(name="微软雅黑", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                          top=Side(style="thin"), bottom=Side(style="thin"))
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # 按PID查找目标行
    found_row = None
    max_row = ws.max_row
    for row_idx in range(2, max_row + 1):
        cell_val = str(ws.cell(row=row_idx, column=pid_col).value or "").strip()
        if cell_val == pid:
            found_row = row_idx
            break

    if found_row is None:
        wb.close()
        return ""

    # 已有的有效字段保留，缺失的字段才用新结果补上
    phone_cell = ws.cell(row=found_row, column=phone_col)
    wechat_cell = ws.cell(row=found_row, column=wechat_col)
    existing_phone = phone_cell.value
    existing_wechat = wechat_cell.value
    if _is_placeholder_contact(existing_phone):
        phone_cell.value = None
    if _is_placeholder_contact(existing_wechat):
        wechat_cell.value = None
    existing_phone_str = _clean_contact_value(existing_phone)
    existing_wechat_str = _clean_contact_value(existing_wechat)

    has_good_phone = bool(existing_phone_str)
    has_good_wechat = bool(existing_wechat_str)

    if has_good_phone and has_good_wechat:
        # 两个字段都已经是有效数据，跳过不覆盖
        wb.close()
        return ""

    final_phone = existing_phone_str if has_good_phone else phone
    final_wechat = existing_wechat_str if has_good_wechat else wechat
    has_real_final = bool((final_phone and final_phone != "无") or final_wechat)

    row_fill = green_fill if has_real_final else red_fill

    # 写入手机号
    ph_cell = ws.cell(row=found_row, column=phone_col)
    ph_cell.value = final_phone
    ph_cell.font = body_font
    ph_cell.border = thin_border
    ph_cell.alignment = center_align
    ph_cell.fill = row_fill

    # 写入微信号
    wx_cell = ws.cell(row=found_row, column=wechat_col)
    wx_cell.value = final_wechat
    wx_cell.font = body_font
    wx_cell.border = thin_border
    wx_cell.alignment = center_align
    wx_cell.fill = row_fill

    # 列宽
    from openpyxl.utils import get_column_letter
    ws.column_dimensions[get_column_letter(phone_col)].width = 16
    ws.column_dimensions[get_column_letter(wechat_col)].width = 22

    wb.save(input_excel)
    wb.close()

    tag = "📱" if has_real_final else "⛔"
    log(f"{tag} [{pid}] → {'phone=' + final_phone if has_real_final else '无（已标记）'}")

    return str(input_path)


# ============================================================
# 命令行入口
# ============================================================
def find_latest_excel() -> str:
    """在采集结果目录找最新的Excel"""
    excels = list(Config.OUTPUT_DIR.glob("颜阿娇_达人_*_*.xlsx"))
    # 排除已含"联系方式"的
    excels = [f for f in excels if "联系方式" not in f.stem]
    if not excels:
        return ""
    excels.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return str(excels[0])


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="颜阿娇 - 批量提取达人联系方式")
    parser.add_argument("--input", type=str, default="",
                         help="输入Excel路径（不指定则自动选最新）")
    parser.add_argument("--max", type=int, default=None,
                         help="最多提取人数（默认50）")
    parser.add_argument("--id", type=str, default="",
                         help="测试单个达人ID")
    parser.add_argument("--delay", type=int, default=Config.PAGE_DELAY,
                         help=f"每人间隔秒数（默认{Config.PAGE_DELAY}）")
    args = parser.parse_args()

    if args.id:
        # 单个测试
        log(f"🔍 测试模式：提取达人 {args.id}")
        p = sync_playwright().start()
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(Config.USER_DATA),
            headless=False,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
            viewport={"width": 1440, "height": 900},
            locale="zh-CN",
        )
        page = ctx.new_page()
        page.goto(Config.DAREN_SQUARE_URL, timeout=60000, wait_until="domcontentloaded")
        time.sleep(5)

        if "login" in page.url.lower():
            log("请在浏览器中扫码登录...")
            for i in range(120):
                time.sleep(2)
                if "login" not in page.url.lower():
                    log("登录成功！")
                    time.sleep(3)
                    break
            else:
                log("登录超时")
                ctx.close()
                p.stop()
                sys.exit(1)

        result = extract_one(page, args.id)
        print(f"\n结果: {json.dumps(result, ensure_ascii=False, indent=2)}")
        ctx.close()
        p.stop()
        sys.exit(0)

    # 批量模式
    Config.PAGE_DELAY = max(Config.MIN_PAGE_DELAY, args.delay)
    input_file = args.input or find_latest_excel()

    if not input_file:
        log("❌ 未找到输入Excel，请用 --input 指定")
        sys.exit(1)

    if not Path(input_file).exists():
        log(f"❌ 文件不存在: {input_file}")
        sys.exit(1)

    batch_extract(input_file, max_count=args.max)