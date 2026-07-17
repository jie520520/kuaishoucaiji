# -*- coding: utf-8 -*-
"""
颜阿娇 - 快手达人采集工具 v10
===============================
独立封装版本：一个脚本搞定全部达人采集。

功能：
  1. 按内容标签精准筛选达人（37个标签全覆盖）
  2. 自动产品契合度评分（内容类型 + 女性受众 + 减肥关联 + 粉丝性价比）
  3. 评分排序 → 条件格式Excel输出
  4. 登录态持久化（一次扫码，长期有效）
  5. 断点续传（已采集标签自动跳过）
  6. 限流自动重试

使用方式：
  python yanajiao_scraper.py --all        # 采集全部37个标签
  python yanajiao_scraper.py --rec        # 采集推荐标签（健康/运动/美妆等）
  python yanajiao_scraper.py --tag 健康    # 采集指定标签（逗号分隔多个）
  python yanajiao_scraper.py              # 交互模式
"""

import os
import sys
import io
import time
import json
import argparse
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        if hasattr(sys.stderr, 'buffer') and not isinstance(sys.stderr, io.TextIOWrapper):
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
        # 关键：强制行缓冲，让输出立即显示到黑窗口
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(line_buffering=True)
    except Exception:
        pass  # 模块导入或管道模式时忽略

# 全局时间标记
_START_TIME = time.time()


def _elapsed() -> str:
    """返回从脚本启动到现在的耗时"""
    secs = int(time.time() - _START_TIME)
    if secs < 60:
        return f"{secs}秒"
    elif secs < 3600:
        return f"{secs//60}分{secs%60}秒"
    return f"{secs//3600}时{(secs%3600)//60}分"


def _ts() -> str:
    """当前时间戳，用于日志前缀"""
    return datetime.now().strftime("%H:%M:%S")


def log(msg: str, end: str = "\n"):
    """带时间戳 + 即刻刷新的打印"""
    text = f"[{_ts()}] {msg}"
    print(text, end=end, flush=True)

from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# ============================================================
# 配置常量
# ============================================================

class Config:
    """全局配置"""
    DAREN_URL = "https://cps.kwaixiaodian.com/zone/daren-match/daren-square-pro"
    API_URL = "https://cps.kwaixiaodian.com/distribute/pc/seller/promoter/list"
    BASE_DIR = Path(__file__).parent
    USER_DATA = BASE_DIR / ".kuaishou_browser_data"
    OUTPUT_DIR = BASE_DIR / "采集结果"
    PAGE_SIZE = 20

    # 全部37个标签
    ALL_TAGS = [
        "三农", "二次元", "亲子", "随手拍", "生活", "穿搭", "美妆", "美食",
        "旅游", "健康", "游戏", "情感", "资讯", "颜值", "运动", "高新数码",
        "动物", "汽车", "音乐", "影视和短剧", "法律", "才艺", "明星娱乐",
        "军事", "教育", "宗教", "读书", "房产家居", "摄影", "舞蹈",
        "搞笑", "财经", "星座命理", "奇人异象", "科学", "历史", "其他",
    ]

    # 推荐标签（与颜阿娇产品匹配度高）
    RECOMMENDED_TAGS = ["健康", "运动", "美妆", "穿搭", "亲子", "生活", "颜值", "舞蹈"]

    # 标签 → API数字ID
    TAG_IDS = {
        "随手拍": [1], "游戏": [2], "二次元": [3], "情感": [4],
        "资讯": [5], "穿搭": [6], "运动": [7], "汽车": [8],
        "音乐": [9], "颜值": [10], "影视和短剧": [11], "亲子": [12],
        "法律": [13], "生活": [14], "才艺": [15], "明星娱乐": [16],
        "军事": [17], "教育": [18], "三农": [19], "宗教": [20],
        "读书": [21], "房产家居": [22], "美妆": [23], "健康": [24],
        "摄影": [25], "动物": [26], "舞蹈": [27], "搞笑": [28],
        "美食": [29], "财经": [30], "星座命理": [31], "高新数码": [32],
        "奇人异象": [33], "旅游": [34], "科学": [35], "历史": [36],
        "其他": [36, 37],
    }

    # 内容标签 → 契合基础分
    FIT_SCORE = {
        "健康": 60, "运动": 50, "美妆": 48, "穿搭": 42,
        "亲子": 40, "生活": 38, "颜值": 35, "舞蹈": 32,
        "情感": 20, "美食": 18, "旅游": 15, "搞笑": 12,
        "才艺": 12, "高新数码": 10, "读书": 10,
    }

    # 带货品类 → 内容标签
    CHANNEL_MAP = {
        "营养健康": "健康", "养生保健": "健康", "运动户外": "运动",
        "美妆护肤": "美妆", "个护清洁": "生活", "女装女鞋": "穿搭",
        "男装男鞋": "穿搭", "母婴玩具": "亲子", "家居百货": "生活",
        "零食饮料": "美食", "生鲜食品": "美食", "茶叶酒水": "美食",
        "珠宝奢品": "穿搭", "数码家电": "高新数码", "图书文娱": "读书",
    }

    # 减肥/体重管理关键词
    WEIGHT_KW = [
        "减肥", "瘦身", "减脂", "燃脂", "身材", "体重", "塑形",
        "体态", "减重", "纤体", "轻体", "排油", "阻断", "控体",
        "掉秤", "燃卡", "刮油", "代餐", "酵素", "低卡", "控糖",
        "代谢", "祛湿", "膳食", "草本", "益生菌",
    ]

    # 女性化昵称特征字
    FEMALE_NICK_CHARS = [
        "妈", "姐", "妹", "娘", "女", "美", "花", "丽", "娜", "芳",
        "娟", "婷", "静", "敏", "雪", "琳", "玲", "菲", "萱", "琪",
    ]


# ============================================================
# 评分引擎
# ============================================================

class Scorer:
    """产品契合度评分（0-100）"""

    @staticmethod
    def score(item: dict, tag: str) -> tuple:
        """
        返回 (总分, 评分明细)
        评分维度：
          - 内容类型匹配 (0-60)
          - 带货品类加成 (0-10)
          - 女性受众匹配 (0-15)
          - 减肥商品关联 (0-15)
          - 粉丝性价比 (0-10)
        """
        parts = []
        total = 0

        # 1. 内容类型 (0-60)
        cs = Config.FIT_SCORE.get(tag, 5)
        total += cs
        parts.append(f"内容:{cs}")

        # 2. 带货品类 (0-10)
        ch = str(item.get("带货品类", ""))
        mapped = Config.CHANNEL_MAP.get(ch, "")
        if mapped == tag:
            cb = 10
        elif mapped in ("健康", "运动", "美妆"):
            cb = 6
        elif mapped in ("穿搭", "生活", "亲子"):
            cb = 4
        elif mapped:
            cb = 2
        else:
            cb = 0
        total += cb
        parts.append(f"品类:{cb}")

        # 3. 女性受众 (0-15)
        gender = str(item.get("达人性别", ""))
        nickname = str(item.get("昵称", ""))
        if gender == "女":
            fs = 15
        elif any(k in nickname for k in Config.FEMALE_NICK_CHARS):
            fs = 8
        else:
            fs = 3
        total += fs
        parts.append(f"女性:{fs}")

        # 4. 减肥商品关联 (0-15)
        items_text = str(item.get("top_items_text", ""))
        ws = min(15, sum(5 for k in Config.WEIGHT_KW if k in items_text))
        total += ws
        parts.append(f"减肥:{ws}")

        # 5. 粉丝性价比 (0-10)
        fans = item.get("粉丝数_raw", 0)
        if 100000 <= fans <= 5000000:
            fsc = 10
        elif 50000 <= fans < 100000:
            fsc = 7
        elif 5000000 < fans <= 10000000:
            fsc = 5
        elif fans > 0:
            fsc = 3
        else:
            fsc = 0
        total += fsc
        parts.append(f"粉丝:{fsc}")

        return total, " + ".join(parts)

    @staticmethod
    def level(score: int) -> str:
        if score >= 65:
            return "⭐⭐⭐ 高度契合"
        if score >= 50:
            return "⭐⭐ 推荐"
        if score >= 35:
            return "⭐ 可考虑"
        return "一般"


# ============================================================
# 核心采集器
# ============================================================

class YanajiaoScraper:
    """
    颜阿娇快手达人采集器
    ── 按内容标签精准筛选 + 产品契合度评分 + 条件格式Excel输出
    """

    def __init__(self):
        self.playwright = None
        self.context = None
        self.page = None
        self._collector_name = "YanajiaoScraper v10"
        self._tag_start_time = 0  # 当前标签开始时间
        # GUI 回调钩子（不设则用模块级 log 函数）
        self._stop = False
        self._on_progress = None       # callback(tag, collected, total, page)
        self._on_tag_done = None       # callback(tag, count, filepath)
        self._on_all_done = None       # callback(results, files, elapsed)
        # 可配置延时（GUI 可调）
        self.page_delay = 2.0          # 页间延迟（秒）
        self.tag_delay = 5             # 标签间延迟（秒）
        self.page_size = Config.PAGE_SIZE  # 每页拉取数量（默认20，可试50）
        # 达人类型筛选（0=全部, 1=直播达人, 2=视频达人）
        self.promoter_type = 0
        # 是否仅采集有联系方式的达人
        self.has_contact = True

    def stop(self):
        """请求停止采集（在当前页完成后生效）"""
        self._stop = True
        log("🛑 收到停止信号，将在当前页完成后停止...")

    # ---------- 浏览器管理 ----------

    def start_browser(self) -> bool:
        """启动浏览器（非 headless，方便扫码登录）"""
        log(f"🚀 启动浏览器...")
        Config.USER_DATA.mkdir(parents=True, exist_ok=True)

        self.playwright = sync_playwright().start()
        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=str(Config.USER_DATA),
            headless=False,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
            viewport={"width": 1440, "height": 900},
            locale="zh-CN",
        )
        self.page = self.context.new_page()
        log("✅ 浏览器已启动")
        return True

    def navigate_to_daren_square(self) -> bool:
        """导航到达人广场页面"""
        log("📡 导航至达人广场...")
        try:
            self.page.goto(Config.DAREN_URL, timeout=60000, wait_until="domcontentloaded")
            time.sleep(3)
            log("✅ 页面加载完成")
            return True
        except Exception as e:
            log(f"⚠️ 导航超时但继续: {e}")
            return True

    def wait_for_login(self, timeout_seconds: int = 180) -> bool:
        """等待用户扫码登录"""
        # 先检查当前是否已登录
        log("🔍 检查登录状态...")
        try:
            body = self.page.locator("body").inner_text(timeout=3000)
            url_lower = self.page.url.lower()
            if "达人广场" in body and "login" not in url_lower:
                log("💚 已有有效登录态，无需扫码")
                time.sleep(2)
                return True
            if "快手" in body and "login" not in url_lower:
                log("🔓 已登录，跳转至达人广场...")
                self.page.goto(Config.DAREN_URL, timeout=30000, wait_until="domcontentloaded")
                time.sleep(3)
                body2 = self.page.locator("body").inner_text(timeout=3000)
                if "达人广场" in body2:
                    log("✅ 已到达达人广场")
                    time.sleep(2)
                    return True
        except Exception:
            pass

        log("")
        log("=" * 50)
        log("  📱 请在浏览器中完成快手扫码登录")
        log("  登录成功后脚本将自动继续")
        log("=" * 50)

        for i in range(timeout_seconds // 2):
            time.sleep(2)
            try:
                url_lower = self.page.url.lower()
                # 已离开登录页 = 登录成功
                if "login" not in url_lower:
                    log("🔓 检测到登录完成，导航至达人广场...")
                    time.sleep(2)
                    self.page.goto(Config.DAREN_URL, timeout=30000, wait_until="domcontentloaded")
                    time.sleep(3)
                    try:
                        body = self.page.locator("body").inner_text(timeout=3000)
                        if "达人广场" in body or "筛选" in body:
                            log("🎉 登录成功！开始采集...")
                            time.sleep(3)
                            return True
                    except Exception:
                        pass
                    if "daren-square" in url_lower or "daren-match" in self.page.url:
                        log("🎉 登录成功！开始采集...")
                        time.sleep(3)
                        return True
            except Exception:
                pass
            if i % 15 == 14:
                log(f"  ⏳ 等待扫码中... (已等 {(i+1)*2} 秒)")

        log("❌ 登录超时，未完成登录")
        return False

    def relogin(self) -> bool:
        """重新登录"""
        log("🔄 需要重新登录，请在浏览器中扫码...")
        self.page.goto(Config.DAREN_URL, timeout=60000, wait_until="domcontentloaded")
        return self.wait_for_login(timeout_seconds=120)

    def close(self):
        """关闭浏览器"""
        log("🔒 关闭浏览器...")
        if self.context:
            try:
                self.context.close()
            except Exception:
                pass
        if self.playwright:
            try:
                self.playwright.stop()
            except Exception:
                pass
        log("👋 浏览器已关闭")

    # ---------- 断点续传 ----------

    @staticmethod
    def get_completed_tags() -> set:
        """扫描已完成的标签（文件名含标签，且文件>5KB）"""
        done = set()
        for f in Config.OUTPUT_DIR.glob("颜阿娇_达人_*_*.xlsx"):
            if f.stat().st_size < 5000:
                continue
            stem = f.stem.replace("颜阿娇_达人_", "")
            # 文件名格式：标签_YYYYMMDD_HHMMSS
            # 时间戳 = 8位日期 + 1下划线 + 6位时间 = 15字符
            if len(stem) > 16 and stem[-16] == '_':
                tag = stem[:-16]
                done.add(tag)
        return done

    # ---------- 数据解析 ----------

    def _parse_item(self, item: dict, tag: str) -> dict:
        """将API返回的单条达人数据解析为标准化字典"""
        vi = item.get("viewInfo") or {}
        nickname = str(item.get("nickname", ""))
        pid = str(item.get("promoterId", ""))

        # 粉丝数
        fans = vi.get("PromoterFansCount") or item.get("fansNum", "")
        if isinstance(fans, (int, float)):
            fans_raw = fans
            fans_display = f"{fans/10000:.1f}万" if fans >= 10000 else str(int(fans))
        else:
            try:
                fans_raw = float(str(fans).replace("万", "")) * 10000
            except Exception:
                fans_raw = 0
            fans_display = str(fans)

        # 带货品类
        hs = item.get("hotSaleChannel", "")
        if isinstance(hs, list):
            hs = "/".join(hs)

        # 热卖商品
        ti = item.get("topItemList", [])
        if isinstance(ti, list):
            tit = " | ".join([
                t.get("itemName", "") for t in ti if t.get("itemName")
            ][:5])
        else:
            tit = ""

        # 佣金率
        comm = item.get("minCommissionRate", "")
        if isinstance(comm, (int, float)):
            comm_display = f"{comm/100:.1f}%"
        else:
            comm_display = str(comm)

        # 客单价
        ap = item.get("promoteAvgPrice", "")
        if isinstance(ap, (int, float)):
            ap_display = f"{ap/100:.2f}"
        else:
            ap_display = str(ap)

        # 性别
        g = item.get("gender", "")
        gd = {"M": "男", "F": "女"}.get(str(g), str(g))

        return {
            "采集时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "内容标签": tag,
            "达人名称": nickname[:50],
            "昵称": nickname,
            "达人性别": gd,
            "快手ID": pid,
            "粉丝数": fans_display,
            "粉丝数_raw": fans_raw,
            "带货品类": str(hs)[:80],
            "top_items_text": tit,
            "top_items_brief": tit[:60] + ("..." if len(tit) > 60 else ""),
            "场均销售额": str(vi.get("PromoterLiveAvgGMV30d", "")),
            "场均观看": str(vi.get("PromoterLiveAvgVisitorNum30d", "")),
            "带货评分": str(vi.get("PromoterScore", "")),
            "佣金率": comm_display,
            "客单价": ap_display,
            "有联系方式": "是",
            "产品契合度": 0,
            "契合等级": "",
            "女性受众": "",
            "评分明细": "",
            "备注": f"佣金率:{comm_display} | 客单价:{ap_display}",
        }

    # ---------- API采集 ----------

    def fetch_tag(self, tag: str, max_retry: int = 8) -> list | None:
        """
        按内容标签拉取该标签下全部达人（有联系方式）。
        返回达人数据列表；登录过期返回 None。
        """
        if tag not in Config.TAG_IDS:
            log(f"  ⛔ 未知标签ID: {tag}，跳过")
            return []

        tag_ids = Config.TAG_IDS[tag]
        params_base = {
            "orderField": 0,
            "orderType": 1,
            "limit": self.page_size,
            "offset": 0,
            "type": self.promoter_type,
            "contentTag": tag_ids,
            "hotSaleChannelIdList": [],
            "hotSaleSubChannelId": [],
        }
        if self.has_contact:
            params_base["hasContact"] = 1

        all_items = []
        offset = 0
        total = None
        page_num = 0
        retry = 0
        self._tag_start_time = time.time()

        log(f"  📥 [{tag}] 开始拉取数据...")
        type_names = {0: "全部达人", 1: "直播达人", 2: "视频达人"}
        log(f"  🔍 类型: {type_names.get(self.promoter_type, '未知')} | "
            f"联系方式: {'仅看有联系方式' if self.has_contact else '不过滤'}")
        # 每3秒发一个心跳点，证明脚本活着
        last_heartbeat = time.time()

        while True:
            # 停止检查
            if self._stop:
                log(f"\n  🛑 [{tag}] 用户停止，已采 {len(all_items)} 条")
                break

            # 心跳：超过3秒没输出就打个点
            if time.time() - last_heartbeat > 3:
                print(".", end="", flush=True)
                last_heartbeat = time.time()

            params = params_base.copy()
            params["offset"] = offset

            try:
                result = self.page.evaluate("""
                    async ({url, body}) => {
                        const r = await fetch(url, {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json',
                                'Accept': 'application/json'
                            },
                            body: JSON.stringify(body),
                            credentials: 'include'
                        });
                        return await r.text();
                    }
                """, {"url": Config.API_URL, "body": params})

                data = json.loads(result)

                if data.get("result") != 1:
                    err_msg = str(data.get("error_msg", ""))
                    if "操作过快" in err_msg or "频繁" in err_msg:
                        retry += 1
                        if retry > max_retry:
                            log(f"\n  ⚠️ 连续限流 {max_retry} 次，已采 {len(all_items)} 条，放弃")
                            break
                        wait = min(60, 3 * (2 ** retry))
                        log(f"\n  ⏳ 限流保护，等待 {wait} 秒 (第{retry}/{max_retry}次)...")
                        time.sleep(wait)
                        last_heartbeat = time.time()
                        continue
                    if "登录" in err_msg or "失效" in err_msg:
                        log(f"\n  🔐 登录过期: {err_msg}")
                        return None
                    log(f"\n  ❌ API错误: {err_msg}")
                    break

                retry = 0
                plist = data.get("data", {}).get("promoterList", [])

                if total is None:
                    total = data.get("data", {}).get("total", 0)
                    log(f"  📊 [{tag}] 共 {total} 人，预计 {max(1, total // self.page_size)} 页")

                if not plist:
                    break

                for item in plist:
                    all_items.append(self._parse_item(item, tag))

                page_num += 1
                offset += self.page_size
                last_heartbeat = time.time()

                # 每页都报进度
                pct = min(100, round(offset / total * 100)) if total else 0
                elapsed = int(time.time() - self._tag_start_time)
                log(f"  📄 [{tag}] 第{page_num}页 | {len(all_items)}/{total}人 ({pct}%) | 耗时{elapsed}秒")

                # GUI 进度回调
                if self._on_progress:
                    try:
                        self._on_progress(tag, len(all_items), total or 0, page_num)
                    except Exception:
                        pass

                if offset >= total:
                    break

                time.sleep(self.page_delay)

            except Exception as e:
                retry += 1
                if retry > max_retry:
                    log(f"\n  ❌ 重试耗尽: {e}")
                    break
                wait = min(30, retry * 5)
                log(f"\n  ⚠️ 异常: {e}，{wait}秒后重试 (第{retry}/{max_retry}次)")
                time.sleep(wait)
                last_heartbeat = time.time()
                continue

        elapsed_tag = int(time.time() - self._tag_start_time)
        log(f"  ✅ [{tag}] 完成！{len(all_items)} 条 | 总耗时 {elapsed_tag} 秒")
        return all_items

    # ---------- 评分 + 保存 ----------

    def score_and_save(self, dlist: list, tag: str) -> Path | None:
        """评分 → 排序 → 保存Excel，返回文件路径"""
        if not dlist:
            log(f"  [{tag}] 无数据，跳过保存")
            return None

        log(f"  📊 [{tag}] 正在评分配对...")
        for d in dlist:
            score, detail = Scorer.score(d, tag)
            d["产品契合度"] = score
            d["契合等级"] = Scorer.level(score)
            d["女性受众"] = (
                "高" if score >= 70 and d.get("达人性别") == "女"
                else ("中" if d.get("达人性别") == "女" else "低")
            )
            d["评分明细"] = detail

        # 按契合度降序 → 粉丝数降序
        dlist.sort(key=lambda x: (-x["产品契合度"], -x.get("粉丝数_raw", 0)))

        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_tag = tag.replace("/", "_").replace("\\", "_")
        filepath = Config.OUTPUT_DIR / f"颜阿娇_达人_{safe_tag}_{timestamp}.xlsx"

        log(f"  💾 [{tag}] 正在保存Excel...")
        wb = Workbook()
        ws = wb.active
        ws.title = f"{tag}达人"

        headers = [
            "序号", "产品契合度", "契合等级", "内容标签", "达人性别", "女性受众",
            "联系方式", "达人名称", "快手ID", "粉丝数", "带货品类",
            "场均销售额", "场均观看", "带货评分", "佣金率", "客单价",
            "商品示例", "评分明细", "采集时间", "备注",
        ]

        # 样式
        hf = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        hfn = Font(color="FFFFFF", bold=True, size=11)
        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        light_green = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        yellow = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid")
        red = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
        border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"),
        )

        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill = hf
            cell.font = hfn
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for ri, d in enumerate(dlist, 2):
            vals = [
                ri - 1, d["产品契合度"], d["契合等级"], d["内容标签"],
                d["达人性别"], d["女性受众"], d["有联系方式"], d["达人名称"],
                d["快手ID"], d["粉丝数"], d["带货品类"], d["场均销售额"],
                d["场均观看"], d["带货评分"], d["佣金率"], d["客单价"],
                d["top_items_brief"], d["评分明细"], d["采集时间"], d["备注"],
            ]
            score = d["产品契合度"]
            fill = (
                green if score >= 65
                else light_green if score >= 50
                else yellow if score >= 35
                else red
            )
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=ri, column=c, value=v)
                cell.fill = fill
                cell.border = border
                cell.alignment = Alignment(vertical="center")

        col_widths = [5, 8, 14, 8, 8, 8, 8, 18, 18, 10, 28, 14, 12, 10, 8, 8, 22, 30, 16, 30]
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

        ws.freeze_panes = "A2"
        Config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        wb.save(filepath)

        hi = sum(1 for d in dlist if d["产品契合度"] >= 65)
        mi = sum(1 for d in dlist if 50 <= d["产品契合度"] < 65)
        fe = sum(1 for d in dlist if d.get("达人性别") == "女")
        log(f"  💾 [{tag}] 已保存: {filepath.name}")
        log(f"  📈 [{tag}] 统计: ≥65分 {hi}人 | 50-64分 {mi}人 | 女性 {fe}人")

        return filepath

    # ---------- 主流程 ----------

    def run(self, tags: list) -> dict:
        """
        主执行流程：登录 → 逐标签采集 → 评分保存 → 汇总
        返回 {tag: count} 的结果字典
        """
        completed = self.get_completed_tags()
        remaining = [t for t in tags if t not in completed]

        log("=" * 60)
        log("  🎯 颜阿娇快手达人采集工具 v10")
        log(f"  📋 全部标签: {len(tags)} 个")
        log(f"  ✅ 已完成: {len(completed)} 个")
        log(f"  📌 待采集: {len(remaining)} 个")
        log("=" * 60)

        if completed:
            log(f"  已完成: {', '.join(sorted(completed))}")
        if not remaining:
            log("🎉 所有标签已采集完毕！")
            return {}

        log(f"  待采集: {', '.join(remaining)}")
        log("")

        # 启动浏览器
        self.start_browser()
        self.navigate_to_daren_square()

        # 登录
        if not self.wait_for_login():
            log("❌ 登录失败，退出")
            self.close()
            return {}

        # 逐标签采集
        results = {}
        all_files = []
        run_start = time.time()

        for idx, tag in enumerate(remaining, 1):
            # 停止检查
            if self._stop:
                log("🛑 用户请求停止，终止后续标签")
                break

            log("")
            log("━" * 50)
            log(f"  📌 [{idx}/{len(remaining)}] 正在处理标签: {tag}")
            log("━" * 50)

            dlist = self.fetch_tag(tag)

            if dlist is None:
                if self.relogin():
                    dlist = self.fetch_tag(tag)
                if dlist is None:
                    log(f"  ⛔ [{tag}] 登录失败，跳过")
                    continue

            fp = self.score_and_save(dlist, tag)
            if fp:
                all_files.append(fp)
                results[tag] = len(dlist)

            # GUI 标签完成回调
            if self._on_tag_done:
                try:
                    self._on_tag_done(tag, len(dlist) if dlist else 0, str(fp) if fp else "")
                except Exception:
                    pass

            # 标签间延迟
            if idx < len(remaining) and not self._stop:
                log(f"  ⏸️ 标签间冷却 {self.tag_delay} 秒...")
                time.sleep(self.tag_delay)

        run_elapsed = int(time.time() - run_start)
        self.close()

        # 汇总
        self._print_summary(tags, completed, results, all_files, run_elapsed)

        # GUI 全部完成回调
        if self._on_all_done:
            try:
                self._on_all_done(results, all_files, run_elapsed)
            except Exception:
                pass

        return results

    def _print_summary(self, all_tags: list, completed: set, results: dict, files: list, elapsed: int):
        """打印采集汇总"""
        log("")
        log("=" * 60)
        log("  📊 采集任务完成！")
        log(f"  ⏱️ 总耗时: {elapsed} 秒")
        log("=" * 60)

        total_count = 0
        for tag in all_tags:
            if tag in completed:
                log(f"  ✅ {tag:<10s} (之前已完成)")
            elif tag in results:
                c = results[tag]
                total_count += c
                log(f"  ⭐ {tag:<10s}: {c:>4d} 人")
            else:
                log(f"  ❌ {tag:<10s}: 失败")

        log(f"\n  📦 本次新增: {len(results)} 个标签, {total_count} 条达人")
        log(f"  📄 生成文件: {len(files)} 个 Excel")
        for f in files:
            log(f"     → {f.name}")
        log("=" * 60)


# ============================================================
# 命令行接口
# ============================================================

def resolve_tags(args) -> list | None:
    """根据命令行参数确定要采集的标签列表"""
    if args.all:
        return Config.ALL_TAGS.copy()
    elif args.rec:
        return Config.RECOMMENDED_TAGS.copy()
    elif args.tag:
        return [t.strip() for t in args.tag.split(",") if t.strip()]
    else:
        return interactive_select()


def interactive_select() -> list | None:
    """交互式选择标签"""
    log("")
    log("=" * 60)
    log("  颜阿娇 - 快手达人采集工具")
    log("=" * 60)

    print("\n📋 可选内容标签：")
    for i, tag in enumerate(Config.ALL_TAGS, 1):
        marker = "⭐" if tag in Config.RECOMMENDED_TAGS else "  "
        print(f"  {marker} {i:2d}. {tag}", end="")
        if i % 4 == 0:
            print()
    print()
    print("⭐ = 推荐标签（与颜阿娇产品匹配度最高）")

    while True:
        print("\n请选择：")
        print("  1) 输入序号（逗号分隔），如: 10,15,7,6")
        print("  2) 输入 'rec' → 采集所有推荐标签")
        print("  3) 输入 'all' → 采集全部标签")
        print("  4) 输入标签名（逗号分隔），如: 健康,运动,美妆")
        print("  0) 退出")

        choice = input("\n> ").strip()
        if not choice:
            continue
        if choice == "0":
            return None
        if choice == "rec":
            return Config.RECOMMENDED_TAGS.copy()
        if choice == "all":
            return Config.ALL_TAGS.copy()

        parts = [p.strip() for p in choice.split(",")]
        result = []
        valid = True
        for part in parts:
            if part.isdigit():
                idx = int(part)
                if 1 <= idx <= len(Config.ALL_TAGS):
                    result.append(Config.ALL_TAGS[idx - 1])
                else:
                    print(f"  无效序号: {idx}")
                    valid = False
                    break
            elif part in Config.ALL_TAGS:
                result.append(part)
            else:
                print(f"  未知标签: {part}")
                valid = False
                break

        if valid and result:
            return result
        print("  请重新输入")


def main():
    parser = argparse.ArgumentParser(
        description="颜阿娇 - 快手达人采集工具 v10",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s --all         采集全部37个标签
  %(prog)s --rec         采集推荐标签
  %(prog)s --tag 健康    采集指定标签
  %(prog)s --tag 健康,运动,美妆  采集多个标签
  %(prog)s               交互模式
        """,
    )
    parser.add_argument("--tag", type=str, help="指定标签名（逗号分隔多个）")
    parser.add_argument("--all", action="store_true", help="采集全部37个标签")
    parser.add_argument("--rec", action="store_true", help="采集推荐标签（健康/运动/美妆/穿搭/亲子/生活/颜值/舞蹈）")
    parser.add_argument("--dry-run", action="store_true", help="预览模式：仅显示将要采集的标签，不实际执行")

    args = parser.parse_args()

    tags = resolve_tags(args)
    if not tags:
        print("未选择任何标签，退出。")
        return

    print(f"\n🎯 将采集 {len(tags)} 个标签：")
    for i, tag in enumerate(tags, 1):
        marker = "⭐" if tag in Config.RECOMMENDED_TAGS else "  "
        print(f"  {marker} {i}. {tag}")

    if args.dry_run:
        print("\n[预览模式] 不会实际执行采集。")
        return

    print()

    scraper = YanajiaoScraper()
    try:
        scraper.run(tags)
    except KeyboardInterrupt:
        log("\n⚠️ 用户中断 (Ctrl+C)")
        scraper.close()
    except Exception as e:
        log(f"\n❌ 致命错误: {e}")
        import traceback
        traceback.print_exc()
        scraper.close()


if __name__ == "__main__":
    main()
