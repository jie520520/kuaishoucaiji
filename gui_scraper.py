# -*- coding: utf-8 -*-
"""
颜阿娇快手达人采集工具 - GUI 版
================================
基于 yanajiao_scraper.py 的图形界面版本。

功能:
  - 可视化选择采集标签 (37个, 推荐标签高亮)
  - 实时进度条 + 彩色日志
  - 可调节采集速度 (页间延迟/标签间延迟/每页数量)
  - 一键启动/停止, 打开输出文件夹
  - 极速模式 (0.5秒/页, 50条/页)

使用:
  双击 启动GUI.bat
  或 python gui_scraper.py
"""

import os
import sys
import io
import time
import queue
import threading
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
        pass

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# 导入采集器模块 (同目录)
sys.path.insert(0, str(Path(__file__).parent))
import yanajiao_scraper as ys
import extract_contacts as ec

# ============================================================
# 颜色 & 字体
# ============================================================

C_BG = "#f0f0f0"
C_LOG_BG = "#1e1e1e"
C_LOG_FG = "#d4d4d4"
C_SUCCESS = "#4ec9b0"
C_WARN = "#dcdcaa"
C_ERROR = "#f44747"
C_INFO = "#569cd6"
C_NORMAL = "#d4d4d4"
C_DIM = "#888888"

F_TITLE = ("Microsoft YaHei UI", 11, "bold")
F_BODY = ("Microsoft YaHei UI", 9)
F_SMALL = ("Microsoft YaHei UI", 8)
F_LOG = ("Consolas", 9)
F_BTN = ("Microsoft YaHei UI", 10, "bold")


# ============================================================
# GUI 主类
# ============================================================

class ScraperGUI:
    """颜阿娇采集工具 GUI"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("颜阿娇 - 快手达人采集工具")
        self.root.geometry("1180x800")
        self.root.minsize(960, 620)
        self.root.configure(bg=C_BG)

        # --- 运行状态 ---
        self.scraper = None
        self.scraper_thread = None
        self.is_running = False
        self.start_time = None
        self.total_collected = 0
        self.files_created = []
        self.completed_tags = 0
        self.total_tags = 0
        self.current_tag = ""

        # --- 线程通信 ---
        self.msg_queue = queue.Queue()

        # --- 控件变量 ---
        self.tag_vars = {}
        self.page_delay_var = tk.StringVar(value="1.0")
        self.tag_delay_var = tk.StringVar(value="3")
        self.page_size_var = tk.StringVar(value="20")
        self.turbo_var = tk.BooleanVar(value=False)
        self.promoter_type_var = tk.StringVar(value="全部达人")
        self.contact_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="* 就绪 - 选择标签后点击开始采集")

        # --- 联系方式提取状态 ---
        self.contact_input_var = tk.StringVar(value="")
        self.contact_delay_var = tk.StringVar(value="10")
        self.contact_max_var = tk.StringVar(value="500")
        self.contact_running = False
        self.contact_scraper_thread = None
        self.contact_start_time = None
        self.contact_extracted = 0
        self.contact_success = 0
        self.contact_skip_pause = False  # 手动跳过暂停标志

        # --- 构建 UI ---
        self._build_ui()
        self._poll_queue()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ========================================================
    # UI 构建
    # ========================================================

    def _build_ui(self):
        # 顶层 Notebook（标签页切换）
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=6, pady=(6, 0))

        # ======== Tab 1: 达人采集 ========
        tab1 = ttk.Frame(nb)
        nb.add(tab1, text=" 达人采集 ")

        left = ttk.LabelFrame(tab1, text="标签选择 & 设置", padding=6)
        left.pack(side="left", fill="y", padx=(0, 4))

        right = ttk.Frame(tab1)
        right.pack(side="right", fill="both", expand=True)

        self._build_left(left)
        self._build_right(right)

        # ======== Tab 2: 联系方式提取 ========
        tab2 = ttk.Frame(nb)
        nb.add(tab2, text=" 联系方式提取 ")
        self._build_contact_tab(tab2)

    def _build_left(self, parent):
        """左侧: 标签选择 + 设置"""

        # --- 标签选择标题 ---
        ttk.Label(parent, text="选择采集标签", font=F_TITLE).pack(anchor="w")
        ttk.Label(parent, text="* = 推荐标签 (与颜阿娇匹配度最高)",
                  font=F_SMALL, foreground=C_DIM).pack(anchor="w", pady=(0, 4))

        # --- 可滚动复选框区域 ---
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill="both", expand=True, pady=(0, 4))

        canvas = tk.Canvas(canvas_frame, width=230, highlightthickness=0, bg=C_BG)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)

        self.scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        win_id = canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _fit_width(e):
            canvas.itemconfig(win_id, width=e.width)
        canvas.bind("<Configure>", _fit_width)

        # 鼠标滚轮 (仅当鼠标在 canvas 区域时生效)
        def _on_wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 创建 37 个复选框
        rec_set = set(ys.Config.RECOMMENDED_TAGS)
        for tag in ys.Config.ALL_TAGS:
            is_rec = tag in rec_set
            prefix = "* " if is_rec else "   "
            var = tk.BooleanVar(value=is_rec)
            self.tag_vars[tag] = var
            cb = ttk.Checkbutton(
                self.scroll_frame,
                text=f"{prefix}{tag}",
                variable=var,
                command=self._update_count,
            )
            cb.pack(anchor="w", padx=4, pady=1)

        # --- 分隔线 ---
        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=6)

        # --- 快捷按钮 ---
        btn_row = ttk.Frame(parent)
        btn_row.pack(fill="x", pady=(0, 4))
        for text, cmd in [("全选", self._sel_all), ("推荐", self._sel_rec),
                          ("清空", self._sel_none), ("反选", self._sel_inv)]:
            ttk.Button(btn_row, text=text, command=cmd).pack(side="left", padx=2)

        self.count_label = ttk.Label(parent, text="已选: 0 个", font=F_BODY)
        self.count_label.pack(anchor="w")
        self._update_count()

        # --- 分隔线 ---
        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=6)

        # --- 采集设置 ---
        sf = ttk.LabelFrame(parent, text="采集设置", padding=6)
        sf.pack(fill="x")

        # 达人类型
        r0 = ttk.Frame(sf)
        r0.pack(fill="x", pady=2)
        ttk.Label(r0, text="达人类型:", font=F_BODY).pack(side="left")
        ttk.Combobox(r0, textvariable=self.promoter_type_var, width=10,
                      values=["全部达人", "直播达人", "视频达人"],
                      state="readonly", font=F_BODY).pack(side="left", padx=4)

        # 联系方式
        r0b = ttk.Frame(sf)
        r0b.pack(fill="x", pady=2)
        ttk.Checkbutton(r0b, text="仅看有联系方式 (不勾选=全部达人)",
                        variable=self.contact_var).pack(side="left")

        # 分隔
        ttk.Separator(sf, orient="horizontal").pack(fill="x", pady=4)

        # 页间延迟
        r1 = ttk.Frame(sf)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="页间延迟:", font=F_BODY).pack(side="left")
        ttk.Spinbox(r1, from_=0.5, to=5.0, increment=0.5,
                     textvariable=self.page_delay_var, width=5,
                     font=F_BODY).pack(side="left", padx=4)
        ttk.Label(r1, text="秒", font=F_SMALL).pack(side="left")

        # 标签间延迟
        r2 = ttk.Frame(sf)
        r2.pack(fill="x", pady=2)
        ttk.Label(r2, text="标签间延迟:", font=F_BODY).pack(side="left")
        ttk.Spinbox(r2, from_=1, to=30, increment=1,
                     textvariable=self.tag_delay_var, width=5,
                     font=F_BODY).pack(side="left", padx=4)
        ttk.Label(r2, text="秒", font=F_SMALL).pack(side="left")

        # 每页数量
        r3 = ttk.Frame(sf)
        r3.pack(fill="x", pady=2)
        ttk.Label(r3, text="每页拉取:", font=F_BODY).pack(side="left")
        ttk.Combobox(r3, textvariable=self.page_size_var, width=5,
                      values=["20", "50"], state="readonly",
                      font=F_BODY).pack(side="left", padx=4)
        ttk.Label(r3, text="条/页", font=F_SMALL).pack(side="left")

        # 极速模式
        ttk.Checkbutton(sf, text="极速模式 (0.5秒/页 + 50条/页)",
                        variable=self.turbo_var,
                        command=self._on_turbo).pack(anchor="w", pady=(6, 0))

        ttk.Label(sf, text="极速模式可能触发限流\n但限流后自动退避重试, 不影响结果",
                  font=F_SMALL, foreground=C_DIM, justify="left").pack(anchor="w", pady=(4, 0))

    def _build_right(self, parent):
        """右侧: 控制按钮 + 进度 + 日志 + 状态栏"""

        # --- 控制按钮 ---
        ctrl = ttk.Frame(parent)
        ctrl.pack(fill="x", pady=(0, 4))

        self.btn_start = ttk.Button(ctrl, text=">>  开始采集", command=self._on_start)
        self.btn_start.pack(side="left", padx=(0, 4))

        self.btn_stop = ttk.Button(ctrl, text="[]  停止", command=self._on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=4)

        ttk.Button(ctrl, text="打开文件夹", command=self._on_open).pack(side="left", padx=4)
        ttk.Button(ctrl, text="清空日志", command=self._clear_log).pack(side="left", padx=4)

        # --- 进度面板 ---
        pf = ttk.LabelFrame(parent, text="采集进度", padding=6)
        pf.pack(fill="x", pady=(0, 4))

        # 当前标签
        r1 = ttk.Frame(pf)
        r1.pack(fill="x", pady=(0, 2))
        self.lbl_cur = ttk.Label(r1, text="当前: -", font=F_BODY)
        self.lbl_cur.pack(side="left")
        self.lbl_cur2 = ttk.Label(r1, text="", font=F_SMALL, foreground=C_DIM)
        self.lbl_cur2.pack(side="right")

        self.pb_tag = ttk.Progressbar(pf, mode="determinate")
        self.pb_tag.pack(fill="x", pady=(0, 6))

        # 总进度
        r2 = ttk.Frame(pf)
        r2.pack(fill="x", pady=(0, 2))
        self.lbl_all = ttk.Label(r2, text="总进度: 0/0", font=F_BODY)
        self.lbl_all.pack(side="left")
        self.lbl_all2 = ttk.Label(r2, text="", font=F_SMALL, foreground=C_DIM)
        self.lbl_all2.pack(side="right")

        self.pb_all = ttk.Progressbar(pf, mode="determinate")
        self.pb_all.pack(fill="x")

        # 统计
        self.lbl_stats = ttk.Label(pf, text="", font=F_SMALL, foreground=C_DIM)
        self.lbl_stats.pack(anchor="w", pady=(4, 0))

        # --- 日志面板 ---
        lf = ttk.LabelFrame(parent, text="运行日志", padding=2)
        lf.pack(fill="both", expand=True, pady=(0, 4))

        self.log_text = tk.Text(
            lf, wrap="word", font=F_LOG,
            bg=C_LOG_BG, fg=C_LOG_FG,
            insertbackground=C_LOG_FG,
            selectbackground="#264f78",
            relief="flat", padx=8, pady=4,
        )
        log_sb = ttk.Scrollbar(lf, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)

        # 日志颜色
        self.log_text.tag_configure("success", foreground=C_SUCCESS)
        self.log_text.tag_configure("warning", foreground=C_WARN)
        self.log_text.tag_configure("error", foreground=C_ERROR)
        self.log_text.tag_configure("info", foreground=C_INFO)
        self.log_text.tag_configure("normal", foreground=C_NORMAL)

        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

        # --- 状态栏 ---
        self.lbl_status = ttk.Label(
            parent, textvariable=self.status_var,
            font=F_BODY, relief="sunken", anchor="w", padding=(8, 3),
        )
        self.lbl_status.pack(fill="x")

    def _build_contact_tab(self, parent):
        """联系方式提取标签页"""
        # 左侧：文件选择 + 设置
        left = ttk.LabelFrame(parent, text="输入 & 设置", padding=6)
        left.pack(side="left", fill="y", padx=(0, 4))

        # 文件选择
        ttk.Label(left, text="选择采集结果Excel", font=F_TITLE).pack(anchor="w")
        ttk.Label(left, text="会提取文件中每个达人的联系方式",
                  font=F_SMALL, foreground=C_DIM).pack(anchor="w", pady=(0, 4))

        fr = ttk.Frame(left)
        fr.pack(fill="x", pady=(0, 4))
        ttk.Entry(fr, textvariable=self.contact_input_var, font=F_BODY,
                  state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(fr, text="浏览", command=self._on_contact_browse
                   ).pack(side="right", padx=(4, 0))

        # 快捷选择
        ttk.Button(left, text="自动选最新Excel", command=self._on_contact_latest
                   ).pack(anchor="w", pady=(0, 6))

        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=4)

        # 设置
        ttk.Label(left, text="提取设置", font=F_TITLE).pack(anchor="w", pady=(4, 0))

        r1 = ttk.Frame(left)
        r1.pack(fill="x", pady=4)
        ttk.Label(r1, text="每人间隔:", font=F_BODY).pack(side="left")
        ttk.Spinbox(r1, from_=10, to=60, increment=1,
                     textvariable=self.contact_delay_var, width=5,
                     font=F_BODY).pack(side="left", padx=4)
        ttk.Label(r1, text="秒", font=F_SMALL).pack(side="left")

        r2 = ttk.Frame(left)
        r2.pack(fill="x", pady=4)
        ttk.Label(r2, text="每次最多提取:", font=F_BODY).pack(side="left")
        ttk.Spinbox(r2, from_=5, to=200, increment=5,
                     textvariable=self.contact_max_var, width=5,
                     font=F_BODY).pack(side="left", padx=4)
        ttk.Label(r2, text="人", font=F_SMALL).pack(side="left")

        ttk.Label(left, text="建议每次不超过50人，避免触发风控",
                  font=F_SMALL, foreground=C_DIM).pack(anchor="w", pady=(0, 4))

        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=8)

        # 注意事项
        ttk.Label(left, text="⚠️ 使用说明", font=F_TITLE).pack(anchor="w")
        notes = [
            "1. 确保已登录快手（与达人采集共用登录态）",
            "2. 浏览器窗口会打开，请勿手动操作",
            "3. 每20人会自动暂停15秒防限流",
            "4. 异常结果当天不重复查询，次日再尝试",
            "5. 解析异常会保存诊断文本，不会写成‘无’",
        ]
        for n in notes:
            ttk.Label(left, text=n, font=F_SMALL,
                      foreground=C_DIM, wraplength=220).pack(anchor="w", pady=1)

        # 右侧：控制 + 进度 + 日志
        right = ttk.Frame(parent)
        right.pack(side="right", fill="both", expand=True)

        # 控制按钮
        ctrl = ttk.Frame(right)
        ctrl.pack(fill="x", pady=(0, 4))

        self.btn_contact_start = ttk.Button(ctrl, text=">>  开始提取联系方式",
                                            command=self._on_contact_start)
        self.btn_contact_start.pack(side="left", padx=(0, 4))

        self.btn_contact_stop = ttk.Button(ctrl, text="[]  停止",
                                           command=self._on_contact_stop,
                                           state="disabled")
        self.btn_contact_stop.pack(side="left", padx=4)

        self.btn_contact_skip = ttk.Button(ctrl, text="⏹  停止暂停",
                                           command=self._on_contact_skip_pause,
                                           state="disabled")
        self.btn_contact_skip.pack(side="left", padx=4)

        ttk.Button(ctrl, text="打开文件夹",
                   command=lambda: os.startfile(str(ys.Config.OUTPUT_DIR))
                   ).pack(side="left", padx=4)

        # 进度
        pf = ttk.LabelFrame(right, text="提取进度", padding=6)
        pf.pack(fill="x", pady=(0, 4))

        self.lbl_contact_progress = ttk.Label(pf, text="等待开始...", font=F_BODY)
        self.lbl_contact_progress.pack(anchor="w")

        self.pb_contact = ttk.Progressbar(pf, mode="determinate")
        self.pb_contact.pack(fill="x", pady=(4, 0))

    def _on_contact_browse(self):
        """浏览选择Excel文件"""
        fp = filedialog.askopenfilename(
            title="选择采集结果Excel",
            initialdir=str(ys.Config.OUTPUT_DIR),
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if fp:
            self.contact_input_var.set(fp)
            self._log(f"已选择: {Path(fp).name}", "info")

    def _on_contact_latest(self):
        """自动选择最新Excel"""
        excels = list(ys.Config.OUTPUT_DIR.glob("颜阿娇_达人_*_*.xlsx"))
        excels = [f for f in excels if "联系方式" not in f.stem]
        if not excels:
            messagebox.showwarning("提示", "采集结果中暂无Excel文件")
            return
        excels.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        self.contact_input_var.set(str(excels[0]))
        self._log(f"自动选择最新: {excels[0].name}", "info")

    def _on_contact_start(self):
        """开始提取联系方式"""
        input_file = self.contact_input_var.get()
        if not input_file:
            messagebox.showwarning("提示", "请先选择采集结果Excel!")
            return
        if not Path(input_file).exists():
            messagebox.showerror("错误", f"文件不存在:\n{input_file}")
            return
        if self.contact_running:
            return

        try:
            delay = max(ec.Config.MIN_PAGE_DELAY, int(self.contact_delay_var.get()))
            self.contact_delay_var.set(str(delay))
            max_count = int(self.contact_max_var.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字!")
            return

        self.contact_running = True
        self.contact_extracted = 0
        self.contact_success = 0
        self.contact_skip_pause = False
        self.contact_start_time = time.time()

        self.btn_contact_start.config(state="disabled")
        self.btn_contact_stop.config(state="normal")
        self.pb_contact["value"] = 0
        self.lbl_contact_progress.config(text="启动浏览器中...")

        self._log("=" * 50, "normal")
        self._log(f">> 开始提取联系方式: {Path(input_file).name}", "info")
        self._log(f"   间隔: {delay}秒 | 上限: {max_count}人", "normal")
        self._log("=" * 50, "normal")

        # 设置参数
        ec.Config.PAGE_DELAY = delay
        ec.Config.MAX_PER_SESSION = max_count

        # 启动后台线程
        self.contact_scraper_thread = threading.Thread(
            target=self._run_contact_extract, args=(input_file, max_count),
            daemon=True
        )
        self.contact_scraper_thread.start()
        self._tick_timer_contact()

    def _on_contact_stop(self):
        """停止提取"""
        self.contact_running = False
        self._log("🛑 收到停止信号...", "warning")
        self.btn_contact_stop.config(state="disabled")
        self.btn_contact_skip.config(state="disabled")

    def _on_contact_skip_pause(self):
        """手动停止暂停，立即恢复提取"""
        self.contact_skip_pause = True
        self._log("⏹  收到停止暂停信号，立即恢复...", "info")
        self.btn_contact_skip.config(state="disabled")

    def _run_contact_extract(self, input_file, max_count):
        """后台线程：批量提取联系方式"""
        from extract_contacts import batch_extract, extract_one, Config as EcConfig
        from playwright.sync_api import sync_playwright

        EcConfig.PAGE_DELAY = max(EcConfig.MIN_PAGE_DELAY, int(self.contact_delay_var.get()))
        EcConfig.MAX_PER_SESSION = max_count

        try:
            results = []

            # 读取Excel
            from openpyxl import load_workbook
            wb = load_workbook(input_file, read_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(min_row=1, values_only=True))
            wb.close()

            if not rows:
                self.msg_queue.put(("contact_error", "Excel为空"))
                return

            header = [str(h).strip() if h else "" for h in rows[0]]
            data_rows = rows[1:]
            ec._clear_placeholder_contacts(input_file)

            # 找列
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
                if h == "手机号":
                    phone_col = i
                if h == "微信号":
                    wechat_col = i

            if pid_col is None:
                self.msg_queue.put(("contact_error", "未找到快手ID列"))
                return

            retry_state = ec._load_retry_state(input_file)
            rows_to_extract = []
            already_has = 0
            retry_exhausted = 0
            for idx, row in enumerate(data_rows):
                pid = str(row[pid_col]).strip() if row[pid_col] else ""
                phone_val, wechat_val = ec._contact_values_from_row(row, phone_col, wechat_col)
                if pid and ec._is_retry_exhausted(retry_state, pid):
                    retry_exhausted += 1
                elif ec._is_contact_complete(retry_state, pid, phone_val, wechat_val):
                    already_has += 1
                else:
                    rows_to_extract.append((idx, row, phone_val, wechat_val))

            total = len(data_rows)
            need_extract = len(rows_to_extract)
            effective = min(max_count, need_extract)

            if already_has > 0:
                self.msg_queue.put(("contact_log", f"   ⏭ 已完成或已可靠分类: {already_has} 人 (自动跳过)"))
            if retry_exhausted > 0:
                self.msg_queue.put(("contact_log", f"   🛡 今日已查询但结果不可靠: {retry_exhausted} 人 (当天不再重复消耗名额)"))
            self.msg_queue.put(("contact_log", f"   🎯 待提取: {need_extract} 人 | 本次: {effective} 人"))

            if need_extract == 0:
                self.msg_queue.put(("contact_log", "✅ 所有达人已完成分类，或今日已达到安全查询上限"))
                self.msg_queue.put(("contact_done", [], input_file))
                return

            self.msg_queue.put(("contact_total", effective))
            self.pb_contact["maximum"] = effective

            # 启动浏览器（辅助函数，GUI内联以共享消息队列）
            def _open_contact_browser():
                self.msg_queue.put(("contact_log", "🚀 启动浏览器..."))
                pp = sync_playwright().start()
                cctx = pp.chromium.launch_persistent_context(
                    user_data_dir=str(EcConfig.USER_DATA),
                    headless=False,
                    args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
                    viewport={"width": 1440, "height": 900},
                    locale="zh-CN",
                )
                ppage = cctx.new_page()
                ppage.goto(EcConfig.DAREN_SQUARE_URL, timeout=60000, wait_until="domcontentloaded")
                time.sleep(5)
                if "login" in ppage.url.lower():
                    self.msg_queue.put(("contact_log", "⚠️  需要扫码登录快手!"))
                return pp, cctx, ppage

            def _close_contact_browser(pp, cctx):
                try:
                    cctx.close()
                except Exception:
                    pass
                try:
                    pp.stop()
                except Exception:
                    pass

            p, ctx, page = _open_contact_browser()

            # 逐个提取（while循环，支持自动暂停恢复）
            consecutive_empty = 0
            pause_count = 0
            stopped_by_user = False
            i = 0
            while i < effective:
                if not self.contact_running:
                    stopped_by_user = True
                    break

                orig_idx, row, existing_phone, existing_wechat = rows_to_extract[i]
                pid = str(row[pid_col]) if row[pid_col] else ""
                name = str(row[name_col]) if name_col is not None and len(row) > name_col and row[name_col] else ""

                if not pid or pid.upper() == "NONE":
                    i += 1
                    continue

                self.msg_queue.put(("contact_progress", i, effective, name, pid))

                result = extract_one(page, pid, name)
                result = ec._record_contact_attempt(
                    input_file, retry_state, pid, existing_phone, existing_wechat, result
                )

                if result["phone"] or result["wechat"]:
                    self.contact_success += 1
                    consecutive_empty = 0
                else:
                    consecutive_empty += 1
                self.contact_extracted += 1
                results.append(result)

                if result["retry_status"] == "解析异常":
                    self.msg_queue.put(("contact_log", f"   🧩 {name or pid}: 页面有字段但解析失败，已保存诊断文本；当天不再查询"))
                elif result["retry_status"] == "平台临时未显示":
                    self.msg_queue.put(("contact_log", f"   🛡 {name or pid}: 平台临时未显示；当天不再查询，明日可重试"))
                elif result["retry_status"] == "技术失败":
                    self.msg_queue.put(("contact_log", f"   🛡 {name or pid}: 技术异常；当天不再自动查询"))
                elif result["retry_status"] == "采集标签与详情冲突":
                    self.msg_queue.put(("contact_log", f"   🛡 {name or pid}: 采集标签确认有联系方式，但详情页未显示；当天不再查询"))
                elif result["retry_status"] in ("仅手机号", "仅微信号"):
                    self.msg_queue.put(("contact_log", f"   ✅ {name or pid}: {result['retry_status']}，页面未展示另一字段"))

                # 保存进度：JSON 每步存（快速可靠），Excel 每5步批量写（防卡锁）
                from extract_contacts import _save_progress, _save_final_excel
                _save_progress(results, input_file)
                if i % 5 == 0 or i == effective - 1 or result["phone"] or result["wechat"]:
                    try:
                        _save_final_excel(results, input_file)
                    except Exception as e:
                        self.msg_queue.put(("contact_log", f"⚠️ Excel保存失败（JSON已存）: {e}"))

                i += 1

                # 检测：连续N个无联系方式 → 自动暂停恢复
                if result.get("rate_limit_warning") or consecutive_empty >= EcConfig.AUTO_PAUSE_CONSECUTIVE_EMPTY:
                    pause_count += 1
                    self.msg_queue.put(("contact_log", ""))
                    self.msg_queue.put(("contact_log", "⚠️" * 20))
                    if result.get("rate_limit_warning"):
                        self.msg_queue.put(("contact_log", f"⚠️  检测到平台提示：{result['rate_limit_warning']}"))
                        self.msg_queue.put(("contact_log", "⚠️  当前达人结果已保存，立即暂停，不再打开下一位达人"))
                    else:
                        self.msg_queue.put(("contact_log", f"⚠️  连续 {consecutive_empty} 个达人未可靠显示联系方式！"))
                        self.msg_queue.put(("contact_log", "⚠️  疑似触达查询限制，触发自动暂停..."))
                    self.msg_queue.put(("contact_log", "⚠️" * 20))

                    # 先保存已有结果
                    _save_final_excel(results, input_file)

                    verification_marker = ec._detect_verification(page)
                    if verification_marker:
                        self.msg_queue.put(("contact_log", f"🧩 检测到‘{verification_marker}’，请在浏览器中人工完成验证"))
                        self.msg_queue.put(("contact_log", "   验证期间不会继续打开达人页面，避免浪费查看名额"))
                        self.lbl_contact_progress.config(text="等待人工完成安全验证...")
                        resolved = False
                        for _ in range(150):
                            if not self.contact_running:
                                stopped_by_user = True
                                break
                            time.sleep(2)
                            if not ec._detect_verification(page):
                                resolved = True
                                break
                        if stopped_by_user:
                            break
                        if not resolved:
                            self.msg_queue.put(("contact_log", "❌ 5分钟内未完成验证，本次任务安全停止"))
                            stopped_by_user = True
                            break
                        self.msg_queue.put(("contact_log", "✅ 验证已完成，继续提取"))
                        consecutive_empty = 0
                        continue

                    # 不关浏览器！保持登录会话存活，避免恢复时需要人工扫码
                    page.goto(EcConfig.DAREN_SQUARE_URL, timeout=30000, wait_until="domcontentloaded")
                    time.sleep(2)

                    wait_minutes = EcConfig.AUTO_PAUSE_MINUTES
                    self.msg_queue.put(("contact_log", f"⏰ 暂停 {wait_minutes} 分钟（浏览器保持打开，无需重新登录）"))
                    self.msg_queue.put(("contact_log", f"   (已提取 {len(results)} 人，成功 {self.contact_success} 人，第 {pause_count} 次暂停)"))
                    self.msg_queue.put(("contact_log", f"   💡 可点击「停止暂停」按钮立即恢复"))
                    self.lbl_contact_progress.config(text=f"⏰ 暂停中... 剩余约 {wait_minutes} 分钟（可点停止暂停）")

                    # 启用跳过暂停按钮
                    self.contact_skip_pause = False
                    self.msg_queue.put(("contact_pause_started",))

                    # 分段等待，每5秒检查一次跳过标志（更快响应）
                    wait_seconds = wait_minutes * 60
                    skipped = False
                    while wait_seconds > 0 and self.contact_running:
                        if self.contact_skip_pause:
                            skipped = True
                            break
                        chunk = min(5, wait_seconds)
                        time.sleep(chunk)
                        wait_seconds -= chunk
                        if wait_seconds > 0:
                            rem_m = wait_seconds // 60
                            rem_s = wait_seconds % 60
                            self.lbl_contact_progress.config(text=f"⏰ 暂停中... 剩余约 {rem_m}分{rem_s}秒（可点停止）")
                            self.msg_queue.put(("contact_log", f"   ⏳ 剩余约 {rem_m} 分 {rem_s} 秒..."))

                    # 禁用跳过暂停按钮
                    self.msg_queue.put(("contact_pause_ended",))

                    if not self.contact_running:
                        stopped_by_user = True
                        break

                    if skipped:
                        self.msg_queue.put(("contact_log", "⏹  用户手动停止暂停，立即恢复提取"))
                    else:
                        self.msg_queue.put(("contact_log", f"⏰ 等待完成，自动恢复"))

                    # 快速验证登录态（导航到广场确认，无需重开浏览器）
                    self.msg_queue.put(("contact_log", ""))
                    self.msg_queue.put(("contact_log", "=" * 40))
                    self.msg_queue.put(("contact_log", f"🔄 第 {pause_count} 次恢复：验证登录态..."))
                    self.lbl_contact_progress.config(text=f"验证登录态...")

                    page.goto(EcConfig.DAREN_SQUARE_URL, timeout=30000, wait_until="domcontentloaded")
                    time.sleep(3)

                    if "login" in page.url.lower():
                        self.msg_queue.put(("contact_log", "⚠️  登录态已过期，等待恢复（30秒自动检测）..."))
                        self.lbl_contact_progress.config(text=f"登录态过期，等待恢复...")
                        logged_in = False
                        for _ in range(15):
                            time.sleep(2)
                            if "login" not in page.url.lower():
                                self.msg_queue.put(("contact_log", "🔓 登录已恢复！"))
                                logged_in = True
                                break
                        if not logged_in:
                            self.msg_queue.put(("contact_log", "⚠️  需要手动扫码，请查看浏览器窗口..."))
                            self.lbl_contact_progress.config(text=f"需要扫码登录（60秒等待）")
                            for _ in range(60):
                                if not self.contact_running:
                                    stopped_by_user = True
                                    break
                                time.sleep(2)
                                if "login" not in page.url.lower():
                                    self.msg_queue.put(("contact_log", "🔓 登录成功！"))
                                    logged_in = True
                                    break
                        if not logged_in:
                            self.msg_queue.put(("contact_log", "❌ 登录超时，停止提取"))
                            stopped_by_user = True
                            break

                    consecutive_empty = 0
                    self.msg_queue.put(("contact_log", f"✅ 已恢复，从第 {i+1}/{effective} 个达人继续"))
                    self.lbl_contact_progress.config(text=f"已恢复，继续提取...")
                    continue

                # 频率控制（正常模式）
                if i < effective and self.contact_running:
                    time.sleep(EcConfig.PAGE_DELAY)
                    if i % EcConfig.LONG_PAUSE_EVERY == 0 and i > 0:
                        self.msg_queue.put(("contact_log", f"  🫁 暂停 {EcConfig.LONG_PAUSE_DURATION} 秒..."))
                        time.sleep(EcConfig.LONG_PAUSE_DURATION)

            # 输出
            if stopped_by_user:
                self.msg_queue.put(("contact_log", ""))
                self.msg_queue.put(("contact_log", "🛑 用户手动停止"))
                self.msg_queue.put(("contact_log", f"   已保存 {self.contact_extracted} 人（成功 {self.contact_success} 人）到原文件"))
                _save_final_excel(results, input_file)
            else:
                output_file = _save_final_excel(results, input_file)
                self.msg_queue.put(("contact_done", results, output_file))

            _close_contact_browser(p, ctx)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.msg_queue.put(("contact_error", str(e)))

    def _tick_timer_contact(self):
        """联系方式提取的计时器"""
        if self.contact_running:
            self.root.after(1000, self._tick_timer_contact)

    # ========================================================
    # 标签选择操作
    # ========================================================

    def _get_tags(self) -> list:
        return [t for t in ys.Config.ALL_TAGS if self.tag_vars[t].get()]

    def _update_count(self):
        n = len(self._get_tags())
        self.count_label.config(text=f"已选: {n} 个")

    def _sel_all(self):
        for v in self.tag_vars.values():
            v.set(True)
        self._update_count()

    def _sel_rec(self):
        rec = set(ys.Config.RECOMMENDED_TAGS)
        for t, v in self.tag_vars.items():
            v.set(t in rec)
        self._update_count()

    def _sel_none(self):
        for v in self.tag_vars.values():
            v.set(False)
        self._update_count()

    def _sel_inv(self):
        for v in self.tag_vars.values():
            v.set(not v.get())
        self._update_count()

    def _on_turbo(self):
        if self.turbo_var.get():
            self.page_delay_var.set("0.5")
            self.tag_delay_var.set("1")
            self.page_size_var.set("50")
            self._log("极速模式已开启: 页间0.5秒, 标签间1秒, 每页50条", "info")
        else:
            self.page_delay_var.set("1.0")
            self.tag_delay_var.set("3")
            self.page_size_var.set("20")
            self._log("极速模式已关闭", "normal")

    # ========================================================
    # 采集控制
    # ========================================================

    def _on_start(self):
        """开始采集"""
        tags = self._get_tags()
        if not tags:
            messagebox.showwarning("提示", "请至少选择一个标签!")
            return
        if self.is_running:
            return

        if len(tags) > 15:
            if not messagebox.askyesno("确认",
                f"将采集 {len(tags)} 个标签, 可能需要较长时间. 继续?"):
                return

        # 重置状态
        self.is_running = True
        self.total_collected = 0
        self.files_created = []
        self.completed_tags = 0
        self.total_tags = len(tags)
        self.start_time = time.time()

        # 更新 UI
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self._set_status("running", f">> 采集中... 0/{self.total_tags} 标签")

        self.pb_all["maximum"] = self.total_tags
        self.pb_all["value"] = 0
        self.pb_tag["value"] = 0
        self.lbl_all.config(text=f"总进度: 0/{self.total_tags}")
        self.lbl_cur.config(text="当前: -")

        self._log("=" * 50, "normal")
        self._log(f">> 开始采集 {len(tags)} 个标签", "info")
        self._log(f"   达人类型: {self.promoter_type_var.get()}", "normal")
        self._log(f"   联系方式: {'仅看有联系方式' if self.contact_var.get() else '全部达人'}", "normal")
        self._log(f"   页间延迟: {self.page_delay_var.get()}秒 | "
                  f"标签间: {self.tag_delay_var.get()}秒 | "
                  f"每页: {self.page_size_var.get()}条", "normal")
        self._log("=" * 50, "normal")

        # 启动采集线程
        self.scraper_thread = threading.Thread(
            target=self._run_scraper, args=(tags,), daemon=True
        )
        self.scraper_thread.start()
        self._tick_timer()

    def _on_stop(self):
        """停止采集"""
        if self.scraper and self.is_running:
            self.scraper.stop()
            self._set_status("stopped", "[] 正在停止...")
            self.btn_stop.config(state="disabled")

    def _run_scraper(self, tags: list):
        """采集线程: 在后台运行, 通过队列与 GUI 通信"""
        # Monkey-patch log 函数
        original_log = ys.log

        def gui_log(msg, end="\n"):
            self.msg_queue.put(("log", msg))

        ys.log = gui_log

        try:
            self.scraper = ys.YanajiaoScraper()
            # 设置回调 -> 通过队列传给 GUI
            self.scraper._on_progress = lambda tag, collected, total, page: \
                self.msg_queue.put(("progress", tag, collected, total, page))
            self.scraper._on_tag_done = lambda tag, count, filepath: \
                self.msg_queue.put(("tag_done", tag, count, filepath))
            self.scraper._on_all_done = lambda results, files, elapsed: \
                self.msg_queue.put(("all_done", results, files, elapsed))

            # 应用设置
            self.scraper.page_delay = float(self.page_delay_var.get())
            self.scraper.tag_delay = int(self.tag_delay_var.get())
            self.scraper.page_size = int(self.page_size_var.get())
            # 达人类型
            type_map = {"全部达人": 0, "直播达人": 1, "视频达人": 2}
            self.scraper.promoter_type = type_map.get(
                self.promoter_type_var.get(), 0)
            # 联系方式
            self.scraper.has_contact = self.contact_var.get()

            # 执行
            self.scraper.run(tags)

        except Exception as e:
            self.msg_queue.put(("error", str(e)))
            import traceback
            traceback.print_exc()
        finally:
            ys.log = original_log
            self.msg_queue.put(("thread_end",))

    def _on_open(self):
        """打开输出文件夹"""
        os.startfile(str(ys.Config.OUTPUT_DIR))

    # ========================================================
    # 队列轮询 -> UI 更新
    # ========================================================

    def _poll_queue(self):
        """每 100ms 轮询一次消息队列"""
        while not self.msg_queue.empty():
            try:
                msg = self.msg_queue.get_nowait()
                mtype = msg[0]

                if mtype == "log":
                    self._log(msg[1])

                elif mtype == "progress":
                    _, tag, collected, total, page = msg
                    self._upd_tag_progress(tag, collected, total, page)

                elif mtype == "tag_done":
                    _, tag, count, filepath = msg
                    self._on_tag_done(tag, count, filepath)

                elif mtype == "all_done":
                    _, results, files, elapsed = msg
                    self._on_all_done(results, files, elapsed)

                elif mtype == "error":
                    self._log(f"致命错误: {msg[1]}", "error")

                elif mtype == "thread_end":
                    self._on_thread_end()

                # --- 联系方式提取消息 ---
                elif mtype == "contact_log":
                    self._log(msg[1])

                elif mtype == "contact_progress":
                    _, idx, total, name, pid = msg
                    self._upd_contact_progress(idx, total, name, pid)

                elif mtype == "contact_total":
                    self.pb_contact["maximum"] = msg[1]

                elif mtype == "contact_done":
                    _, results, output_file = msg
                    self._on_contact_done(results, output_file)

                elif mtype == "contact_error":
                    self._log(f"❌ 提取错误: {msg[1]}", "error")
                    self.contact_running = False
                    self.btn_contact_start.config(state="normal")
                    self.btn_contact_stop.config(state="disabled")
                    self.btn_contact_skip.config(state="disabled")

                elif mtype == "contact_pause_started":
                    self.btn_contact_skip.config(state="normal")

                elif mtype == "contact_pause_ended":
                    self.btn_contact_skip.config(state="disabled")

            except queue.Empty:
                break
        self.root.after(100, self._poll_queue)

    # ========================================================
    # UI 更新方法
    # ========================================================

    def _log(self, msg: str, tag: str = None):
        """追加日志 (自动着色 + 时间戳 + 自动滚动)"""
        if tag is None:
            # 根据 emoji 自动判断颜色
            tag = self._detect_log_color(msg)

        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {msg}\n", tag)
        self.log_text.see("end")

        # 限制日志行数 (保留最后 1500 行)
        lines = int(self.log_text.index("end-1c").split(".")[0])
        if lines > 2000:
            self.log_text.delete("1.0", f"{lines - 1500}.0")

    @staticmethod
    def _detect_log_color(msg: str) -> str:
        """根据消息内容中的 emoji 判断日志颜色"""
        # 成功: 绿色
        for e in ["\u2705", "\U0001f49a", "\U0001f389", "\U0001f680",
                   "\u2713", "\U0001f513"]:
            if e in msg:
                return "success"
        # 警告: 黄色
        for e in ["\u26a0", "\u23f3", "\u23f8", "\u23f1"]:
            if e in msg:
                return "warning"
        # 错误: 红色
        for e in ["\u274c", "\u26d4", "\U0001f510", "\U0001f6d1"]:
            if e in msg:
                return "error"
        # 信息: 蓝色
        for e in ["\U0001f4cc", "\U0001f4e5", "\U0001f4ca", "\U0001f4be",
                   "\U0001f4c4", "\U0001f4c8", "\U0001f4e1", "\U0001f50d",
                   "\U0001f50c"]:
            if e in msg:
                return "info"
        return "normal"

    def _upd_tag_progress(self, tag, collected, total, page):
        """更新当前标签的进度条"""
        self.current_tag = tag
        pct = min(100, int(collected / total * 100)) if total else 0
        self.pb_tag["value"] = pct

        self.lbl_cur.config(
            text=f"当前: {tag} ({self.completed_tags + 1}/{self.total_tags})"
        )
        self.lbl_cur2.config(text=f"{collected}/{total} ({pct}%) | 第{page}页")
        self._upd_stats()

    def _on_tag_done(self, tag, count, filepath):
        """一个标签完成"""
        self.completed_tags += 1
        self.total_collected += count
        if filepath:
            self.files_created.append(filepath)
        self.pb_all["value"] = self.completed_tags
        self.lbl_all.config(text=f"总进度: {self.completed_tags}/{self.total_tags}")
        self.lbl_all2.config(text=f"已采集 {self.total_collected} 人")
        self._upd_stats()

    def _on_all_done(self, results, files, elapsed):
        """全部完成"""
        total = sum(results.values())
        self._log("", "normal")
        self._log("=" * 50, "success")
        self._log("采集任务完成!", "success")
        self._log(f"   总耗时: {self._fmt(elapsed)}", "normal")
        self._log(f"   新增: {len(results)} 标签, {total} 条达人", "normal")
        self._log(f"   文件: {len(files)} 个 Excel", "normal")
        self._log("=" * 50, "success")
        self._set_status("done",
            f"OK 完成 | {total}人 | {len(files)}文件 | {self._fmt(elapsed)}")
        self.pb_tag["value"] = 100

    def _on_thread_end(self):
        """采集线程结束"""
        self.is_running = False
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        current = self.status_var.get()
        if "完成" not in current and "错误" not in current and "停止" not in current:
            self._set_status("ready", "* 就绪")

    def _upd_stats(self):
        """更新统计行"""
        elapsed = int(time.time() - self.start_time) if self.start_time else 0
        files = len(self.files_created)
        self.lbl_stats.config(
            text=f"已采集: {self.total_collected} 人 | "
                 f"文件: {files} 个 | 耗时: {self._fmt(elapsed)}"
        )

    def _tick_timer(self):
        """每秒刷新一次耗时显示"""
        if self.is_running:
            self._upd_stats()
            self.root.after(1000, self._tick_timer)

    def _set_status(self, level, text):
        """更新状态栏"""
        colors = {
            "ready": C_DIM,
            "running": "#00aa00",
            "stopped": "#cc6600",
            "done": "#0066cc",
            "error": "#cc0000",
        }
        self.status_var.set(text)
        self.lbl_status.config(foreground=colors.get(level, C_DIM))

    def _clear_log(self):
        """清空日志"""
        self.log_text.delete("1.0", "end")

    # ========================================================
    # 联系方式提取 UI 更新
    # ========================================================

    def _upd_contact_progress(self, idx, total, name, pid):
        """更新联系方式提取进度"""
        pct = int((idx + 1) / total * 100)
        self.pb_contact["value"] = idx + 1
        self.lbl_contact_progress.config(
            text=f"[{idx+1}/{total}] {name} (ID:{pid}) | "
                 f"已提取{self.contact_extracted}人, 成功{self.contact_success}人"
        )

    def _on_contact_done(self, results, output_file):
        """联系方式提取完成"""
        self.contact_running = False
        total = len(results)
        success = sum(1 for r in results if r.get("phone") or r.get("wechat"))
        elapsed = int(time.time() - self.contact_start_time) if self.contact_start_time else 0

        self._log("", "normal")
        self._log("=" * 50, "success")
        self._log("联系方式提取完成!", "success")
        self._log(f"   总人数: {total} | 提取成功: {success} | "
                  f"耗时: {self._fmt(elapsed)}", "normal")
        self._log(f"   输出: {Path(output_file).name}", "normal")
        self._log("=" * 50, "success")

        self.lbl_contact_progress.config(
            text=f"✅ 完成! {success}/{total} 人成功 | 耗时 {self._fmt(elapsed)}"
        )
        self.pb_contact["value"] = total
        self.btn_contact_start.config(state="normal")
        self.btn_contact_stop.config(state="disabled")
        self.btn_contact_skip.config(state="disabled")

    # ========================================================
    # 工具方法
    # ========================================================

    @staticmethod
    def _fmt(secs):
        if secs < 60:
            return f"{secs}秒"
        elif secs < 3600:
            return f"{secs//60}分{secs%60}秒"
        return f"{secs//3600}时{(secs%3600)//60}分"

    def _on_close(self):
        """窗口关闭"""
        if self.is_running:
            if not messagebox.askyesno("确认", "采集正在进行中, 确定退出?"):
                return
            if self.scraper:
                self.scraper.stop()
                time.sleep(1)
        self.root.destroy()

    # ========================================================
    # 启动
    # ========================================================

    def run(self):
        self._log("欢迎使用颜阿娇快手达人采集工具", "info")
        self._log("选择左侧标签后点击 [开始采集]", "normal")
        self._log("首次使用需扫码登录快手, 之后自动保持登录态", "normal")
        self._log("", "normal")
        self.root.mainloop()


# ============================================================
# 入口
# ============================================================

if __name__ == "__main__":
    app = ScraperGUI()
    app.run()
