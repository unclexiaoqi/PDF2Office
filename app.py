"""PDF 转 Word — desktop app (tkinter)."""

from __future__ import annotations

import json
import os
import queue
import sys
import tempfile
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
import uuid
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    HAS_DND = True
except Exception:  # pragma: no cover - optional dependency
    HAS_DND = False

from converter import (
    CONVERSIONS,
    CONVERSION_ORDER,
    DEFAULT_CONVERSION,
    BatchConverter,
    FileResult,
    human_size,
)

APP_TITLE = "PDF2Office"
APP_VERSION = "2.1.2"
SETTINGS_VERSION = 2

# ---------------------------------------------------------------------------
# 首次启动的免责声明确认（参照同类客户端的“复述声明”做法，输入一段话才能用）
# ---------------------------------------------------------------------------
DISCLAIMER_VERSION = "2026-11-25"
DISCLAIMER_SUMMARY = """PDF2Office 是一款本地格式转换工具，仅供学习、研究与技术交流使用，不是商业产品。

· 只转换你有合法权利处理的文件，并遵守所在地法律法规；
· 转换全部在本机完成：不联网、不上传文件、不收集任何数据；
· 不提供破解加密、绕过 DRM 或规避授权限制的功能；
· 软件按「现状」提供，不担保转换结果完全无误，重要文件请先备份并人工复核；
· 因使用或无法使用本软件产生的后果由你自行承担，开发者不承担责任。

完整条款见项目主页的 DISCLAIMER.md。"""
CONSENT_PHRASE = "我只转换我有权处理的文件，并遵守所在地法律法规，自行承担使用后果。"
CLI_DISCLAIMER_LINE = "PDF2Office：仅供学习与技术交流，请只处理你拥有合法权利的文件；完整条款见 DISCLAIMER.md。"

BG = "#F4F6FB"
WHITE = "#FFFFFF"
TEXT = "#1E293B"
MUTED = "#6B7A90"
FAINT = "#94A3B8"
ACCENT = "#2563EB"
ACCENT_DARK = "#1D4ED8"
ACCENT_SOFT = "#EFF6FF"
ACCENT_LINE = "#BFDBFE"
BORDER = "#E3E8F0"
GREEN = "#15803D"
GREEN_SOFT = "#F0FDF4"
RED = "#DC2626"
RED_SOFT = "#FEF2F2"

FONT_FAMILY = "Microsoft YaHei UI"
OUTPUT_PLACEHOLDER = "与 PDF 文件放在一起"


def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def config_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / "PDF2Office"


# ---------------------------------------------------------------------------
# 免责声明的接受状态（存在 settings.json 里，随版本号失效重问）
# ---------------------------------------------------------------------------
def _norm_consent(text: str) -> str:
    """只比“有效字符”：忽略空格、换行与中英文标点差异。"""
    return "".join(ch for ch in text if ch.isalnum())


def _consent_matches(typed: str) -> bool:
    return _norm_consent(typed) == _norm_consent(CONSENT_PHRASE)


def _read_settings_dict() -> dict:
    try:
        return json.loads((config_dir() / "settings.json").read_text("utf-8"))
    except Exception:
        return {}


def disclaimer_accepted() -> bool:
    return _read_settings_dict().get("disclaimer") == DISCLAIMER_VERSION


def record_disclaimer_acceptance() -> None:
    try:
        folder = config_dir()
        folder.mkdir(parents=True, exist_ok=True)
        data = _read_settings_dict()
        data["disclaimer"] = DISCLAIMER_VERSION
        (folder / "settings.json").write_text(
            json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
    except Exception:
        pass


def round_rect(canvas: tk.Canvas, x1, y1, x2, y2, r, **kwargs):
    pts = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return canvas.create_polygon(pts, smooth=True, **kwargs)


class ModeCard(tk.Frame):
    def __init__(self, parent, key: str, preset: dict, on_select, fonts,
                 compact=False, slim=False):
        super().__init__(parent, bg=WHITE, highlightthickness=1,
                         highlightbackground=BORDER, cursor="hand2")
        self.key = key
        self.selected = False
        self.on_select = on_select
        self.preset = preset

        self.pad = tk.Frame(self, bg=WHITE)
        self.pad.pack(fill="both", expand=True, padx=12, pady=5 if (compact or slim) else 9)

        self.top = tk.Frame(self.pad, bg=WHITE)
        self.top.pack(fill="x")
        self.dot = tk.Label(self.top, text="○", bg=WHITE, fg=FAINT,
                            font=(FONT_FAMILY, 13), cursor="hand2")
        self.dot.pack(side="left")
        self.name = tk.Label(self.top, text=self.preset["name"], bg=WHITE, fg=TEXT,
                             font=fonts["card"], cursor="hand2")
        self.name.pack(side="left", padx=(7, 0))
        self.badge = tk.Label(self.top, text=self.preset["short"], bg="#F1F5F9", fg=MUTED,
                              font=fonts["tiny"], padx=7, pady=1, cursor="hand2")
        self.badge.pack(side="right")

        self.desc = tk.Label(self.pad, text=self.preset["desc"], bg=WHITE, fg=MUTED,
                             font=fonts["small"], anchor="w", cursor="hand2")
        self.desc.pack(fill="x", padx=(24, 0), pady=(1, 0))
        self.tag = tk.Label(self.pad, text=self.preset.get("tag", ""), bg=WHITE, fg=FAINT,
                            font=fonts["tiny"], anchor="w", cursor="hand2")
        if not compact and not slim:
            self.tag.pack(fill="x", padx=(24, 0))

        widgets = [self, self.pad, self.top, self.dot, self.name, self.badge, self.desc]
        if not slim:
            widgets.append(self.tag)
        for w in widgets:
            w.bind("<Button-1>", self._click)
            w.bind("<Enter>", self._enter)
            w.bind("<Leave>", self._leave)

        badge = self.preset.get("badge")
        if badge:
            self.badge.configure(text=f"{badge} · {self.preset['short']}",
                                 bg=ACCENT_SOFT, fg=ACCENT)

    def _click(self, _event=None):
        self.on_select(self.key)

    def _enter(self, _event=None):
        if not self.selected:
            self._paint("#F8FAFC", BORDER)

    def _leave(self, _event=None):
        self._paint(WHITE if not self.selected else ACCENT_SOFT,
                    ACCENT if self.selected else BORDER)

    def _paint(self, bg: str, border: str):
        self.configure(bg=bg, highlightbackground=border,
                       highlightthickness=2 if self.selected else 1)
        for w in (self.pad, self.top, self.dot, self.name, self.desc, self.tag):
            w.configure(bg=bg)

    def set_selected(self, value: bool):
        self.selected = value
        self.dot.configure(text="●" if value else "○", fg=ACCENT if value else FAINT)
        self._paint(ACCENT_SOFT if value else WHITE, ACCENT if value else BORDER)
        if value:
            self.configure(highlightthickness=2)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} · 本地转换工具")
        self.root.configure(bg=BG)

        self.screen_w = self.root.winfo_screenwidth()
        self.screen_h = self.root.winfo_screenheight()
        self.compact = self.screen_h < 1000 or os.environ.get("PDF2OFFICE_COMPACT") == "1"
        win_w = 1000 if self.compact else 1080
        win_h = min(700, self.screen_h - 80) if self.compact else min(820, self.screen_h - 90)
        pos_x = max(0, (self.screen_w - win_w) // 2)
        pos_y = max(0, (self.screen_h - win_h) // 2 - 24)
        self.root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.root.minsize(960, 660)

        icon = resource_path(os.path.join("assets", "icon.ico"))
        if os.path.exists(icon):
            try:
                self.root.iconbitmap(default=icon)
            except Exception:
                pass

        self.files: list[str] = []
        self.conversion = DEFAULT_CONVERSION
        self.modes = {name: spec["default_mode"] for name, spec in CONVERSIONS.items()}
        self.mode = self.modes[self.conversion]
        self.output_mode = "same"
        self.output_dir = ""
        self.overwrite = False
        self.last_dir = str(Path.home())
        self.disclaimer = ""
        self.proc = None
        self.worker = None
        self.queue = None
        self.cancel_event = threading.Event()
        self.running = False
        self.started_at = 0.0
        self.timer_job = None
        self.errors: dict[str, str] = {}
        self.run_files: list[str] = []
        self.last_outdir = None
        self.progress_value = 0

        self._init_fonts()
        self._init_styles()
        self._load_settings()
        self._build_ui()
        self._refresh_list()
        self._update_mode_cards()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Control-o>", lambda e: self.add_files())

    # ------------------------------------------------------------------ setup
    def _init_fonts(self):
        global FONT_FAMILY
        families = set(tkfont.families(self.root))
        family = FONT_FAMILY
        for cand in ("Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", "Segoe UI"):
            if cand in families:
                family = cand
                break
        FONT_FAMILY = family
        self.fonts = {
            "title": (family, 17 if self.compact else 19, "bold"),
            "subtitle": (family, 9),
            "section": (family, 10, "bold"),
            "card": (family, 10, "bold"),
            "body": (family, 9),
            "small": (family, 8),
            "tiny": (family, 8),
            "button": (family, 9, "bold"),
            "ghost": (family, 9),
            "file": (family, 9),
            "status": (family, 10, "bold"),
            "hint": (family, 8),
        }

    def _init_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Files.Treeview", background=WHITE, fieldbackground=WHITE,
                        foreground=TEXT, borderwidth=0, rowheight=34,
                        font=self.fonts["file"])
        style.configure("Files.Treeview.Heading", background="#F8FAFC", foreground=MUTED,
                        relief="flat", borderwidth=0, font=self.fonts["tiny"])
        style.map("Files.Treeview.Heading", background=[("active", "#F1F5F9")])
        style.map("Files.Treeview", background=[("selected", ACCENT_SOFT)],
                  foreground=[("selected", TEXT)])
        style.layout("Files.Treeview", [("Files.Treeview.treearea", {"sticky": "nswe"})])
        style.configure("Accent.Horizontal.TProgressbar",
                        troughcolor="#E9EEF7", background=ACCENT,
                        bordercolor="#DCE4F0", lightcolor=ACCENT,
                        darkcolor=ACCENT, borderwidth=1, thickness=12)
        style.configure("Vertical.TScrollbar", background="#D7DEEA", troughcolor=WHITE,
                        borderwidth=0, arrowsize=10, width=10)

    def _load_settings(self):
        path = config_dir() / "settings.json"
        try:
            data = json.loads(path.read_text("utf-8"))
        except Exception:
            return
        upgraded = False
        try:
            old_version = int(data.get("settings_version", 0))
        except (TypeError, ValueError):
            old_version = 0
        if old_version < SETTINGS_VERSION:
            # v1 was the PDF→Word-only build; its mode names do not all exist
            # in the merged app, so start from clean defaults once.
            upgraded = True
        else:
            saved = data.get("conversion")
            if saved in CONVERSIONS:
                self.conversion = saved
            stored = data.get("modes") or {}
            for name, spec in CONVERSIONS.items():
                value = stored.get(name)
                if value in spec["modes"]:
                    self.modes[name] = value
            single = data.get("mode")
            if single and single in CONVERSIONS[self.conversion]["modes"]:
                self.modes[self.conversion] = single
        self.mode = self.modes[self.conversion]
        self.output_mode = data.get("output_mode", "same")
        self.output_dir = data.get("output_dir", "")
        self.overwrite = bool(data.get("overwrite", False))
        self.last_dir = data.get("last_dir", self.last_dir)
        self.disclaimer = data.get("disclaimer", "")
        if upgraded:
            self._save_settings()

    def _save_settings(self):
        try:
            d = config_dir()
            d.mkdir(parents=True, exist_ok=True)
            self.modes[self.conversion] = self.mode
            (d / "settings.json").write_text(json.dumps({
                "settings_version": SETTINGS_VERSION,
                "conversion": self.conversion,
                "mode": self.mode,
                "modes": self.modes,
                "output_mode": self.output_mode,
                "output_dir": self.output_dir,
                "overwrite": self.overwrite,
                "last_dir": self.last_dir,
                "disclaimer": self.disclaimer,
            }, ensure_ascii=False, indent=2), "utf-8")
        except Exception:
            pass

    # --------------------------------------------------------------------- UI
    def _build_ui(self):
        self._build_header()
        self._build_bottom()

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=18 if self.compact else 22,
                  pady=(12, 8) if self.compact else (14, 10))
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, minsize=352)
        body.rowconfigure(0, weight=1)

        self._build_files(body)
        self._build_modes(body)

    def _build_header(self):
        header = tk.Frame(self.root, bg=WHITE, height=84 if self.compact else 92)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Frame(header, bg=ACCENT, height=3).pack(fill="x", side="top")

        inner = tk.Frame(header, bg=WHITE)
        inner.pack(fill="both", expand=True, padx=22, pady=12)

        icon = tk.Canvas(inner, width=46, height=46, bg=WHITE, highlightthickness=0)
        icon.pack(side="left", pady=2)
        round_rect(icon, 2, 2, 44, 44, 11, fill=ACCENT, outline="")
        icon.create_text(23, 24, text="PDF", fill="white", font=(FONT_FAMILY, 9, "bold"))

        titles = tk.Frame(inner, bg=WHITE)
        titles.pack(side="left", padx=12)
        tk.Label(titles, text="PDF2Office", bg=WHITE, fg=TEXT,
                 font=self.fonts["title"]).pack(anchor="w")
        tk.Label(titles, text=f"本地转换 · 保留原版式 · PDF ↔ Word / PPT  v{APP_VERSION}",
                 bg=WHITE, fg=MUTED, font=self.fonts["subtitle"]).pack(anchor="w")

        badge = tk.Frame(inner, bg=GREEN_SOFT, highlightthickness=1,
                         highlightbackground="#BBF7D0")
        badge.pack(side="right", pady=10)
        tk.Label(badge, text="●", bg=GREEN_SOFT, fg=GREEN,
                 font=(FONT_FAMILY, 7)).pack(side="left", padx=(9, 2))
        tk.Label(badge, text="本地处理 · 文件不上传", bg=GREEN_SOFT, fg=GREEN,
                 font=self.fonts["tiny"]).pack(side="left", padx=(0, 9), pady=3)

    def _build_files(self, parent):
        card = tk.Frame(parent, bg=WHITE, highlightthickness=1,
                        highlightbackground=BORDER)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        card.rowconfigure(3, weight=1)
        card.columnconfigure(0, weight=1)

        head = tk.Frame(card, bg=WHITE)
        head.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 0))
        tk.Label(head, text="待转换文件", bg=WHITE, fg=TEXT,
                 font=self.fonts["section"]).pack(side="left")
        self.count_label = tk.Label(head, text="0 个", bg=WHITE, fg=FAINT,
                                    font=self.fonts["tiny"])
        self.count_label.pack(side="left", padx=(6, 0))

        self.btn_add = self._ghost_button(head, "＋ 添加文件", self.add_files)
        self.btn_add.pack(side="right")
        self.btn_add_dir = self._ghost_button(head, "添加文件夹", self.add_folder)
        self.btn_add_dir.pack(side="right", padx=(0, 6))

        self.files_hint = tk.Label(card, text=CONVERSIONS[self.conversion]["hint"],
                                   bg=WHITE, fg=FAINT, font=self.fonts["hint"], anchor="w")
        self.files_hint.grid(row=1, column=0, sticky="ew", padx=16, pady=(4, 8))

        tree_wrap = tk.Frame(card, bg=WHITE)
        tree_wrap.grid(row=3, column=0, sticky="nsew", padx=(16, 8), pady=(0, 6))
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_wrap, style="Files.Treeview", show="headings",
                                 columns=("name", "size", "status"), selectmode="extended")
        self.tree.heading("name", text="文件名", anchor="w")
        self.tree.heading("size", text="大小", anchor="e")
        self.tree.heading("status", text="状态", anchor="w")
        self.tree.column("name", width=310, minwidth=180, stretch=True, anchor="w")
        self.tree.column("size", width=76, minwidth=66, stretch=False, anchor="e")
        self.tree.column("status", width=138, minwidth=96, stretch=False, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.tag_configure("odd", background="#FAFBFF")
        self.tree.tag_configure("ok", foreground=TEXT)
        self.tree.tag_configure("err", foreground=RED)
        self.tree.tag_configure("busy", foreground=ACCENT)
        self.tree.bind("<Double-1>", self._on_tree_double)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        scroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview,
                               style="Vertical.TScrollbar")
        scroll.grid(row=0, column=1, sticky="ns", padx=(2, 0))
        self.tree.configure(yscrollcommand=scroll.set)

        self.empty_hint = tk.Frame(tree_wrap, bg=WHITE)
        tk.Label(self.empty_hint, text="⇩", bg=WHITE, fg="#CBD5E1",
                 font=(FONT_FAMILY, 30)).pack()
        tk.Label(self.empty_hint, text="拖入文件，或点击「添加文件」", bg=WHITE,
                 fg=MUTED, font=self.fonts["body"]).pack(pady=(2, 0))
        self.empty_hint.place(relx=0.5, rely=0.46, anchor="center")

        foot = tk.Frame(card, bg=WHITE)
        foot.grid(row=4, column=0, sticky="ew", padx=16, pady=(0, 12))
        self.btn_remove = self._ghost_button(foot, "移除选中", self.remove_selected, fg=MUTED)
        self.btn_remove.pack(side="left")
        self.btn_clear = self._ghost_button(foot, "清空", self.clear_files, fg=MUTED)
        self.btn_clear.pack(side="left", padx=(4, 0))
        tk.Label(foot, text="转换后可双击文件行查看错误详情", bg=WHITE, fg=FAINT,
                 font=self.fonts["hint"]).pack(side="right")

        if HAS_DND and hasattr(self.root, "drop_target_register"):
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _build_modes(self, parent):
        panel = tk.Frame(parent, bg=BG)
        panel.grid(row=0, column=1, sticky="nsew")
        panel.columnconfigure(0, weight=1)

        tk.Label(panel, text="转换类型", bg=BG, fg=TEXT,
                 font=self.fonts["section"]).grid(row=0, column=0, sticky="w")

        selector = tk.Frame(panel, bg=BG)
        selector.grid(row=1, column=0, sticky="ew", pady=(4, 9))
        self.conv_buttons = {}
        for key in CONVERSION_ORDER:
            button = tk.Button(
                selector, text=CONVERSIONS[key]["short"],
                command=lambda name=key: self._select_conversion(name),
                relief="flat", bd=0, highlightthickness=1,
                highlightbackground=BORDER, bg=WHITE, fg=MUTED,
                activebackground="#F1F5F9", activeforeground=ACCENT,
                font=self.fonts["tiny"], padx=4, pady=6, cursor="hand2",
            )
            button.pack(side="left", expand=True, fill="x", padx=(0, 5))
            self.conv_buttons[key] = button

        tk.Label(panel, text="输出方式", bg=BG, fg=TEXT,
                 font=self.fonts["section"]).grid(row=2, column=0, sticky="w")
        self.mode_hint = tk.Label(panel, text="", bg=BG, fg=FAINT,
                                  font=self.fonts["hint"])
        self.mode_hint.grid(row=3, column=0, sticky="w", pady=(2, 6))

        self.cards_wrap = tk.Frame(panel, bg=BG)
        self.cards_wrap.grid(row=4, column=0, sticky="ew")
        self.cards = {}
        self._rebuild_mode_cards()

        tk.Label(panel, text="输出位置", bg=BG, fg=TEXT,
                 font=self.fonts["section"]).grid(row=5, column=0, sticky="w", pady=(8, 4))

        self.output_var = tk.StringVar(value=self.output_mode)
        radios = tk.Frame(panel, bg=BG)
        radios.grid(row=6, column=0, sticky="w")
        self.radios = []
        for value, label in (("same", "与源文件相同目录"), ("custom", "指定文件夹")):
            rb = tk.Radiobutton(radios, text=label, variable=self.output_var, value=value,
                                command=self._on_output_mode, bg=BG, fg=TEXT,
                                activebackground=BG, selectcolor=WHITE,
                                font=self.fonts["body"], highlightthickness=0,
                                bd=0, cursor="hand2")
            rb.pack(side="left", padx=(0, 14))
            self.radios.append(rb)

        row = tk.Frame(panel, bg=BG)
        row.grid(row=7, column=0, sticky="ew", pady=(4, 0))
        row.columnconfigure(0, weight=1)
        self.out_entry = tk.Entry(row, bg=WHITE, fg=TEXT, relief="flat",
                                  highlightthickness=1, highlightbackground=BORDER,
                                  highlightcolor=ACCENT, insertbackground=TEXT,
                                  font=self.fonts["body"], disabledbackground="#F1F5F9",
                                  disabledforeground=FAINT)
        self.out_entry.grid(row=0, column=0, sticky="ew", ipady=5, padx=(0, 6))
        if self.output_dir:
            self.out_entry.insert(0, self.output_dir)
        self.btn_browse = self._ghost_button(row, "浏览", self.choose_output_dir)
        self.btn_browse.grid(row=0, column=1)

        self.overwrite_var = tk.BooleanVar(value=self.overwrite)
        self.overwrite_check = tk.Checkbutton(
            panel, text="覆盖同名文件（默认自动重命名）",
            variable=self.overwrite_var, command=self._on_overwrite,
            bg=BG, fg=MUTED, activebackground=BG, selectcolor=WHITE,
            font=self.fonts["small"], highlightthickness=0, bd=0,
            cursor="hand2")
        self.overwrite_check.grid(row=8, column=0, sticky="w", pady=(5, 0))

        self._on_output_mode(save=False)
        self._update_conversion_buttons()

    def _rebuild_mode_cards(self):
        for card in self.cards.values():
            card.destroy()
        self.cards = {}
        spec = CONVERSIONS[self.conversion]
        for key in spec["order"]:
            card = ModeCard(self.cards_wrap, key, spec["modes"][key],
                            self._select_mode, self.fonts,
                            compact=self.compact, slim=True)
            card.pack(fill="x", pady=(0, 5))
            self.cards[key] = card
        self.mode_hint.configure(text=self._mode_hint())
        self._update_mode_cards()

    def _mode_hint(self) -> str:
        if self.conversion == "ppt2pdf":
            return "「原图无损」按原始像素写图，文字仍是矢量"
        if self.conversion == "pdf2word":
            return "「原版式」版式 100% 一致；「可编辑」文字可改"
        return "「原版式」版式 100% 一致；「可编辑」文字拆成文本框"

    def _update_conversion_buttons(self):
        for key, button in self.conv_buttons.items():
            if key == self.conversion:
                button.configure(bg=ACCENT, fg="white", highlightbackground=ACCENT,
                                 activebackground=ACCENT_DARK,
                                 activeforeground="white")
            else:
                button.configure(bg=WHITE, fg=MUTED, highlightbackground=BORDER,
                                 activebackground="#F1F5F9",
                                 activeforeground=ACCENT)

    def _select_conversion(self, key: str):
        if self.running or key == self.conversion:
            return
        self.modes[self.conversion] = self.mode
        self.conversion = key
        self.mode = self.modes.get(key) or CONVERSIONS[key]["default_mode"]
        accept = CONVERSIONS[key]["accept"]
        kept = [p for p in self.files if Path(p).suffix.lower() in accept]
        dropped = len(self.files) - len(kept)
        self.files = kept
        self._update_conversion_buttons()
        self._rebuild_mode_cards()
        self._refresh_list()
        self._save_settings()
        text = f"已切换到 {CONVERSIONS[key]['name']}"
        if dropped:
            text += f"，移除了 {dropped} 个格式不匹配的文件"
        elif kept:
            text += f"，{len(kept)} 个文件可继续使用"
        self._set_status(text)

    def _build_bottom(self):
        bar = tk.Frame(self.root, bg=WHITE, height=80 if self.compact else 86,
                       highlightthickness=1, highlightbackground=BORDER)
        bar.pack(side="bottom", fill="x")
        bar.pack_propagate(False)
        inner = tk.Frame(bar, bg=WHITE)
        inner.pack(fill="both", expand=True, padx=22, pady=12)
        inner.columnconfigure(0, weight=1)

        self.status_label = tk.Label(inner, text="准备就绪 · 添加 PDF 文件开始转换",
                                     bg=WHITE, fg=TEXT, font=self.fonts["status"], anchor="w")
        self.status_label.grid(row=0, column=0, sticky="ew")

        prog = tk.Frame(inner, bg=WHITE)
        prog.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        prog.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(prog, style="Accent.Horizontal.TProgressbar",
                                        mode="determinate", maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew")
        self.percent_label = tk.Label(prog, text="0%", bg=WHITE, fg=TEXT,
                                      font=self.fonts["small"], width=5, anchor="e")
        self.percent_label.grid(row=0, column=1, padx=(10, 0))
        self.timer_label = tk.Label(prog, text="", bg=WHITE, fg=FAINT,
                                    font=self.fonts["hint"], width=9, anchor="e")
        self.timer_label.grid(row=0, column=2, padx=(6, 0))

        actions = tk.Frame(inner, bg=WHITE)
        actions.grid(row=0, column=1, rowspan=2, sticky="e", padx=(18, 0))
        self.btn_open = self._ghost_button(actions, "打开输出文件夹", self.open_output, fg=MUTED)
        self.btn_open.pack(side="left", padx=(0, 8))
        self.btn_open.configure(state="disabled")
        self.btn_stop = self._ghost_button(actions, "停止", self.stop_convert, fg=RED)
        self.btn_stop.configure(state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 8))
        self.btn_convert = self._primary_button(actions, "开始转换", self.start_convert)
        self.btn_convert.pack(side="left")

    # ------------------------------------------------------------------ widgets
    def _ghost_button(self, parent, text, command, fg=ACCENT):
        btn = tk.Button(parent, text=text, command=command, bg=WHITE, fg=fg,
                        activebackground="#F1F5F9", activeforeground=fg,
                        relief="flat", bd=0, highlightthickness=0,
                        font=self.fonts["ghost"], cursor="hand2",
                        padx=9, pady=5)
        btn.bind("<Enter>", lambda e: btn.configure(bg="#F1F5F9"))
        btn.bind("<Leave>", lambda e: btn.configure(bg=WHITE))
        return btn

    def _primary_button(self, parent, text, command):
        return tk.Button(parent, text=text, command=command, bg=ACCENT, fg="white",
                         activebackground=ACCENT_DARK, activeforeground="white",
                         disabledforeground="#DBEAFE",
                         relief="flat", bd=0, highlightthickness=0,
                         font=self.fonts["button"], cursor="hand2",
                         padx=22, pady=9)

    # ------------------------------------------------------------------ files
    def _accept(self) -> set[str]:
        return CONVERSIONS[self.conversion]["accept"]

    def _on_drop(self, event):
        if self.running:
            return
        try:
            paths = self.root.tk.splitlist(event.data)
        except Exception:
            paths = [event.data]
        self._add_paths([str(p) for p in paths])

    def add_files(self):
        if self.running:
            return
        if self.conversion == "ppt2pdf":
            patterns = ("PowerPoint 演示文稿",
                        "*.ppt *.pptx *.pptm *.pps *.ppsx *.potx *.potm")
            title = "选择 PPT 文件"
        else:
            patterns = ("PDF 文件", "*.pdf")
            title = "选择 PDF 文件"
        paths = filedialog.askopenfilenames(
            parent=self.root, title=title, initialdir=self.last_dir,
            filetypes=[patterns, ("所有文件", "*.*")],
        )
        if paths:
            self.last_dir = os.path.dirname(paths[0])
            self._add_paths(list(paths))

    def add_folder(self):
        if self.running:
            return
        kind = "PPT" if self.conversion == "ppt2pdf" else "PDF"
        folder = filedialog.askdirectory(parent=self.root, title=f"选择包含 {kind} 的文件夹",
                                         initialdir=self.last_dir)
        if not folder:
            return
        self.last_dir = folder
        accept = self._accept()
        found = []
        for p in Path(folder).rglob("*"):
            if p.is_file() and p.suffix.lower() in accept and not p.name.startswith("~$"):
                found.append(str(p))
        if not found:
            messagebox.showinfo(APP_TITLE, f"这个文件夹里没有找到 {kind} 文件。",
                                parent=self.root)
            return
        self._add_paths(found)

    def _add_paths(self, paths):
        existing = {os.path.normcase(os.path.abspath(p)) for p in self.files}
        accept = self._accept()
        added = 0
        skipped = 0
        for p in paths:
            if not os.path.isfile(p):
                continue
            if Path(p).suffix.lower() not in accept:
                skipped += 1
                continue
            key = os.path.normcase(os.path.abspath(p))
            if key in existing:
                continue
            existing.add(key)
            self.files.append(os.path.abspath(p))
            added += 1
        if added:
            self._refresh_list()
            self._set_status(f"已添加 {added} 个文件，共 {len(self.files)} 个")
        elif skipped:
            self._set_status(f"这些文件的格式不适合「{CONVERSIONS[self.conversion]['name']}」")
        else:
            self._set_status("没有新增文件（可能已存在或格式不支持）")

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for i, path in enumerate(self.files):
            try:
                size = human_size(os.path.getsize(path))
            except OSError:
                size = "—"
            tags = ["odd"] if i % 2 else []
            self.tree.insert("", "end", iid=path,
                             values=(os.path.basename(path), size, "等待转换"),
                             tags=tags)
        self.count_label.configure(text=f"{len(self.files)} 个")
        if hasattr(self, "files_hint"):
            self.files_hint.configure(text=CONVERSIONS[self.conversion]["hint"])
        if self.files:
            self.empty_hint.place_forget()
        else:
            self.empty_hint.place(relx=0.5, rely=0.46, anchor="center")
        self.btn_convert.configure(state="normal" if self.files and not self.running else "disabled")

    def remove_selected(self):
        if self.running:
            return
        selected = list(self.tree.selection())
        if not selected:
            return
        drop = set(selected)
        self.files = [p for p in self.files if p not in drop]
        self._refresh_list()

    def clear_files(self):
        if self.running or not self.files:
            return
        self.files.clear()
        self._refresh_list()
        self._set_status("已清空文件列表")

    def _on_tree_double(self, _event):
        selected = self.tree.selection()
        if not selected:
            return
        path = selected[0]
        err = self.errors.get(path)
        if err:
            self._show_error_dialog(path, err)

    # ------------------------------------------------------------------ modes
    def _select_mode(self, key):
        if self.running:
            return
        self.mode = key
        self.modes[self.conversion] = key
        self._update_mode_cards()
        self._save_settings()

    def _update_mode_cards(self):
        for key, card in self.cards.items():
            card.set_selected(key == self.mode)

    def _on_output_mode(self, save=True):
        if self.running:
            return
        mode = self.output_var.get()
        if mode == "custom":
            self.out_entry.configure(state="normal")
            if self.out_entry.get().strip() in ("", OUTPUT_PLACEHOLDER):
                self.out_entry.delete(0, "end")
                if self.output_dir:
                    self.out_entry.insert(0, self.output_dir)
            self.btn_browse.configure(state="normal")
        else:
            text = self.out_entry.get().strip()
            if text and text != OUTPUT_PLACEHOLDER:
                self.output_dir = text
            self.out_entry.configure(state="normal")
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, OUTPUT_PLACEHOLDER)
            self.out_entry.configure(state="disabled")
            self.btn_browse.configure(state="disabled")
        self.output_mode = mode
        if save:
            self._save_settings()

    def choose_output_dir(self):
        if self.running:
            return
        folder = filedialog.askdirectory(parent=self.root, title="选择输出文件夹",
                                         initialdir=self.out_entry.get() or self.last_dir)
        if folder:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, folder)
            self.output_dir = folder
            self._save_settings()

    def _on_overwrite(self):
        self.overwrite = bool(self.overwrite_var.get())
        self._save_settings()

    # ------------------------------------------------------------------ run
    def start_convert(self):
        if self.running:
            return
        if not self.files:
            kind = "PPT" if self.conversion == "ppt2pdf" else "PDF"
            messagebox.showinfo(APP_TITLE, f"请先添加要转换的 {kind} 文件。", parent=self.root)
            return
        custom = self.output_var.get() == "custom"
        outdir = None
        if custom:
            outdir = self.out_entry.get().strip()
            if outdir == OUTPUT_PLACEHOLDER:
                outdir = ""
            if not outdir:
                messagebox.showwarning(APP_TITLE, "请选择输出文件夹，或改为“与源文件相同目录”。",
                                       parent=self.root)
                return
            if not os.path.isdir(outdir):
                try:
                    os.makedirs(outdir, exist_ok=True)
                except Exception as exc:
                    messagebox.showerror(APP_TITLE, f"无法创建输出文件夹：\n{exc}",
                                         parent=self.root)
                    return
            self.output_dir = outdir

        self.errors.clear()
        self.last_outdir = None
        self.run_files = list(self.files)
        self.cancel_event = threading.Event()
        self.queue = queue.Queue()
        self.running = True
        self.started_at = time.time()

        for path in self.run_files:
            self.tree.set(path, "status", "等待中")
            self.tree.item(path, tags=["odd"] if self.files.index(path) % 2 else [])

        conv = BatchConverter(
            self.run_files, self.conversion, self.mode, outdir,
            bool(self.overwrite_var.get()),
            self.cancel_event, lambda kind, payload: self.queue.put((kind, payload)),
        )
        self.worker = threading.Thread(target=conv.run, daemon=True)
        self.worker.start()

        self._set_running(True)
        self._set_status("正在准备转换引擎…")
        self.progress.configure(mode="indeterminate")
        try:
            self.progress.start(12)
        except Exception:
            pass
        self._tick_timer()
        self._poll_queue()

    def stop_convert(self):
        if not self.running:
            return
        self.cancel_event.set()
        self.btn_stop.configure(state="disabled", text="正在停止…")
        self._set_status("将在当前文件完成后停止…")

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                self._handle_event(kind, payload)
        except queue.Empty:
            pass
        if self.running or not self.queue.empty():
            self.root.after(80, self._poll_queue)

    def _handle_event(self, kind, payload):
        if kind == "stage":
            self._set_status(payload.get("text", ""))
            return

        if kind == "start":
            idx = payload["index"]
            name = payload["name"]
            path = self.run_files[idx]
            total = max(1, len(self.run_files))
            # The render/export step gives no sub-progress, so show a moving
            # bar rather than a dead one; real numbers replace it as they arrive.
            try:
                self.progress.stop()
            except Exception:
                pass
            self.progress.configure(mode="indeterminate", maximum=100)
            try:
                self.progress.start(14)
            except Exception:
                pass
            self.percent_label.configure(text=f"{int(idx / total * 100)}%")
            self.tree.set(path, "status", "转换中…")
            self.tree.item(path, tags=["busy"])
            self.tree.see(path)
            self._set_status(f"正在转换 {idx + 1}/{len(self.run_files)}：{name}")
            return

        if kind == "progress":
            idx = int(payload.get("index", 0))
            fraction = max(0.0, min(1.0, float(payload.get("fraction") or 0.0)))
            total = max(1, len(self.run_files))
            self._show_progress((idx + fraction) / total)
            return

        if kind == "compress":
            idx = payload["index"]
            path = self.run_files[idx]
            self.tree.set(path, "status", payload.get("text") or "正在转换…")
            return

        if kind == "file_cancelled":
            idx = payload["index"]
            path = self.run_files[idx]
            self.tree.set(path, "status", "已停止")
            self.tree.item(path, tags=[])
            self._show_progress((idx + 1) / max(1, len(self.run_files)))
            return

        if kind in ("file_ok", "file_error"):
            result = payload["result"]
            idx = payload["index"]
            path = self.run_files[idx]
            note = getattr(result, "note", "") or ""
            if result.ok:
                text = f"✓ 完成 · {human_size(result.size_after)}"
                if note:
                    text += f"（{note}）"
                self.tree.set(path, "status", text)
                self.tree.item(path, tags=["ok"])
                self.last_outdir = os.path.dirname(result.dst) if result.dst else None
                self.btn_open.configure(state="normal")
            elif getattr(result, "cancelled", False):
                self.tree.set(path, "status", "已停止")
                self.tree.item(path, tags=[])
            else:
                self.tree.set(path, "status", "✗ 失败")
                self.tree.item(path, tags=["err"])
                self.errors[path] = (note + " " + (result.error or "未知错误")).strip()
                self._record_error(path, self.errors[path])
            self._show_progress((idx + 1) / max(1, len(self.run_files)))
            return

        if kind == "done":
            self._finish(payload)
            return

        if kind == "fatal":
            self._set_running(False)
            self.progress.stop()
            self.progress.configure(mode="determinate", maximum=100, value=0)
            self.percent_label.configure(text="0%")
            self.timer_label.configure(text="")
            self._set_status("转换失败")
            self._record_error("(batch)", payload.get("text", "未知错误"))
            messagebox.showerror(APP_TITLE, payload.get("text", "未知错误"), parent=self.root)

    def _show_progress(self, overall: float):
        """Drive the bottom bar as a determinate progress bar (0..1)."""
        value = max(0.0, min(1.0, overall)) * 100.0
        try:
            if str(self.progress.cget("mode")) != "determinate":
                self.progress.stop()
                self.progress.configure(mode="determinate", maximum=100)
            self.progress.configure(value=value)
        except Exception:
            pass
        self.percent_label.configure(text=f"{int(round(value))}%")

    def _set_running(self, running: bool):
        self.running = running
        state = "disabled" if running else "normal"
        for btn in (self.btn_add, self.btn_add_dir, self.btn_remove, self.btn_clear,
                    self.btn_convert):
            btn.configure(state=state)
        for button in getattr(self, "conv_buttons", {}).values():
            button.configure(state=state)
        self.btn_stop.configure(state="normal" if running else "disabled")
        if not running:
            self.btn_stop.configure(text="停止")
        for rb in getattr(self, "radios", []):
            rb.configure(state="disabled" if running else "normal")
        if hasattr(self, "overwrite_check"):
            self.overwrite_check.configure(state="disabled" if running else "normal")
        self.out_entry.configure(state="disabled" if running else
                                 ("normal" if self.output_mode == "custom" else "disabled"))
        self.btn_browse.configure(state="disabled" if running else
                                  ("normal" if self.output_mode == "custom" else "disabled"))
        self.btn_convert.configure(text="转换中…" if running else "开始转换")
        if running:
            self.btn_convert.configure(state="disabled")
        else:
            self.btn_convert.configure(state="normal" if self.files else "disabled")

    def _tick_timer(self):
        if not self.running:
            self.timer_label.configure(text="")
            return
        elapsed = int(time.time() - self.started_at)
        self.timer_label.configure(text=f"已用时 {elapsed // 60:02d}:{elapsed % 60:02d}")
        self.timer_job = self.root.after(500, self._tick_timer)

    def _finish(self, payload):
        results = payload.get("results", [])
        cancelled = payload.get("cancelled", False)
        engine = (payload.get("engine") or "").strip()
        self._set_running(False)
        ok = [r for r in results if r.ok]
        failed = [r for r in results if not r.ok and not r.cancelled]
        self._show_progress(1.0 if results else 0.0)

        if cancelled:
            self._set_status(f"已停止 · 完成 {len(ok)} 个")
            return
        if not results:
            self._set_status("没有文件被转换")
            return

        saved = sum(max(0, r.size_before - r.size_after) for r in ok)
        grown = sum(max(0, r.size_after - r.size_before) for r in ok)
        total_before = sum(r.size_before for r in ok)
        parts = [f"完成 {len(ok)} 个"]
        if failed:
            parts.append(f"失败 {len(failed)} 个（双击失败行查看原因）")
        if total_before > 0 and saved > grown:
            ratio = saved / total_before * 100.0
            if ratio >= 1.0:
                parts.append(f"体积减少 {human_size(saved)}（{ratio:.0f}%）")
        elif grown > total_before * 0.01:
            parts.append(f"输出 {human_size(sum(r.size_after for r in ok))}")
        if engine and "wps" in engine.lower():
            parts.append("WPS 引擎")
        self._set_status(" · ".join(parts))
        self.timer_label.configure(text="")
        if failed and not ok:
            detail = "\n\n".join(
                f"{Path(r.src).name}\n{r.error or '未知错误'}" for r in failed[:3]
            )
            messagebox.showerror("转换失败", detail, parent=self.root)
        if failed:
            self.btn_open.configure(state="normal" if self.last_outdir else "disabled")

    def _record_error(self, path: str, error: str):
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        text = f"[{stamp}] {path}\n{error}\n\n"
        for folder in (config_dir(), Path(os.environ.get("TEMP", "."))):
            try:
                folder.mkdir(parents=True, exist_ok=True)
                with (folder / "pdf2office-errors.log").open("a", encoding="utf-8") as fh:
                    fh.write(text)
            except Exception:
                pass

    def _set_status(self, text: str):
        self.status_label.configure(text=text)

    # ------------------------------------------------------------------ misc
    def open_output(self):
        folder = self.last_outdir
        if not folder or not os.path.isdir(folder):
            messagebox.showinfo(APP_TITLE, "还没有生成 PDF 文件。", parent=self.root)
            return
        try:
            os.startfile(folder)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"无法打开文件夹：\n{exc}", parent=self.root)

    def _show_error_dialog(self, path, error):
        win = tk.Toplevel(self.root)
        win.title("转换失败详情")
        win.configure(bg=WHITE)
        win.geometry("640x360")
        win.transient(self.root)
        tk.Label(win, text=os.path.basename(path), bg=WHITE, fg=TEXT,
                 font=self.fonts["section"], anchor="w").pack(fill="x", padx=16, pady=(14, 2))
        tk.Label(win, text=path, bg=WHITE, fg=FAINT, font=self.fonts["hint"],
                 anchor="w", wraplength=600, justify="left").pack(fill="x", padx=16)
        text = tk.Text(win, bg="#F8FAFC", fg=TEXT, relief="flat", font=self.fonts["body"],
                       wrap="word", padx=10, pady=10)
        text.pack(fill="both", expand=True, padx=16, pady=10)
        text.insert("1.0", error)
        text.configure(state="disabled")
        self._primary_button(win, "知道了", win.destroy).pack(anchor="e", padx=16, pady=(0, 14))

    def _on_close(self):
        if self.running:
            if not messagebox.askyesno(APP_TITLE, "正在转换中，确定要退出吗？", parent=self.root):
                return
            self.cancel_event.set()
        self._save_settings()
        try:
            self.root.destroy()
        except Exception:
            pass


def _result_dict(result: FileResult) -> dict:
    return {
        "src": result.src,
        "dst": result.dst,
        "ok": bool(result.ok),
        "error": result.error,
        "pages": int(result.pages or 0),
        "size_before": int(result.size_before or 0),
        "size_after": int(result.size_after or 0),
        "seconds": float(result.seconds or 0),
        "cancelled": bool(result.cancelled),
        "note": getattr(result, "note", "") or "",
    }


def _cli(argv):
    """Headless batch mode used both directly and by the GUI helper."""
    outdir = ""
    conversion = DEFAULT_CONVERSION
    mode = "" 
    state_path = ""
    cancel_path = ""
    files = []
    index = 0
    while index < len(argv):
        arg = argv[index]
        if arg in ("--out", "--mode", "--conversion", "--state", "--cancel") and index + 1 < len(argv):
            value = argv[index + 1]
            if arg == "--out":
                outdir = value
            elif arg == "--mode":
                mode = value
            elif arg == "--conversion":
                conversion = value
            elif arg == "--state":
                state_path = value
            else:
                cancel_path = value
            index += 2
            continue
        if arg != "--convert" and not arg.startswith("--"):
            files.append(arg)
        index += 1

    files = [os.path.abspath(f) for f in files if os.path.isfile(f)]
    if outdir:
        outdir = os.path.abspath(outdir)
    if conversion not in CONVERSIONS:
        conversion = DEFAULT_CONVERSION
    spec = CONVERSIONS[conversion]
    if mode not in spec["modes"]:
        mode = spec["default_mode"]

    state = {
        "stage": "starting", "done": False, "fatal": "", "cancelled": False,
        "index": 0, "total": len(files), "name": "", "results": [],
    }
    state_lock = threading.Lock()

    def write_state():
        if not state_path:
            return
        try:
            with state_lock:
                tmp = state_path + ".tmp"
                with open(tmp, "w", encoding="utf-8") as fh:
                    json.dump(state, fh, ensure_ascii=False)
                os.replace(tmp, state_path)
        except Exception:
            pass

    if not files or mode not in spec["modes"]:
        state["done"] = True
        state["fatal"] = "参数错误：没有可转换的文件，或输出方式无效。"
        write_state()
        return 2

    cancel_event = threading.Event()
    if cancel_path:
        def watch_cancel():
            while not cancel_event.is_set():
                if os.path.exists(cancel_path):
                    cancel_event.set()
                    return
                time.sleep(0.2)

        threading.Thread(target=watch_cancel, daemon=True).start()

    def log(text):
        try:
            if sys.stderr is not None:
                print(text, file=sys.stderr)
        except Exception:
            pass

    def emit(kind, payload):
        if kind == "stage":
            state["stage"] = "starting"
            write_state()
        elif kind == "start":
            state["stage"] = "converting"
            state["index"] = int(payload.get("index", 0))
            state["name"] = payload.get("name", "")
            write_state()
        elif kind == "compress":
            state["stage"] = "compress"
            write_state()
        elif kind in ("file_ok", "file_error", "file_cancelled"):
            result = payload["result"]
            state["results"].append(_result_dict(result))
            state["stage"] = "converting"
            state["index"] = int(payload.get("index", 0)) + 1
            write_state()
        elif kind == "done":
            state["done"] = True
            state["stage"] = "done"
            state["cancelled"] = bool(payload.get("cancelled", False))
            state["results"] = [_result_dict(r) for r in payload.get("results", [])]
            write_state()
            for item in payload.get("results", []):
                log(f"{'OK ' if item.ok else 'ERR'} {item.src} -> {item.dst or ''} {item.error or ''}")
        elif kind == "fatal":
            state["done"] = True
            state["fatal"] = payload.get("text", "未知错误")
            state["stage"] = "fatal"
            write_state()
            log("[fatal] " + state["fatal"])

    if state_path:
        write_state()
    converter = BatchConverter(files, conversion, mode, outdir or None, True,
                               cancel_event, emit)
    converter.run()
    if not state["done"]:
        state["done"] = True
        state["fatal"] = state.get("fatal") or "转换进程意外结束。"
        write_state()
    results = state["results"]
    return 0 if results and all(item.get("ok") for item in results) else 1


def _ensure_writable_temp():
    """Pick a writable directory for this run's own scratch files.

    Only Python's tempfiles are redirected — the TEMP/TMP of child processes
    (cscript and the Office engine) are deliberately left alone.  Redirecting
    them breaks Office add-ins, which keep their state under %TEMP%, and can
    make the engine fail to open documents.
    """
    candidates = []
    normal = os.environ.get("TEMP") or os.environ.get("TMP")
    if normal:
        candidates.append(Path(normal) / "PDF2Office-work")
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidates.append(Path(local) / "PDF2Office" / "work")
    profile = os.environ.get("USERPROFILE")
    if profile:
        candidates.append(Path(profile) / "PDF2Office-work")
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).parent / "tmp")
    candidates.append(Path(os.environ.get("SystemDrive", "C:") + r"\Temp\PDF2Office-work"))
    for folder in candidates:
        try:
            folder.mkdir(parents=True, exist_ok=True)
            probe = folder / f"write-test-{uuid.uuid4().hex}.txt"
            probe.write_text("ok", "utf-8")
            probe.unlink(missing_ok=True)
            tempfile.tempdir = str(folder)
            return folder
        except Exception:
            continue
    return None


def show_disclaimer(parent) -> bool:
    """使用前的确认框：读完声明、完整输入那句承诺，才能点「同意」。

    参照同类客户端的做法：不是打个勾，而是要求手动输入一遍，确保用户真的看过。
    """
    dialog = tk.Toplevel(parent)
    dialog.title(f"{APP_TITLE} · 免责声明")
    dialog.configure(bg=BG)
    dialog.resizable(False, False)
    width, height = 720, 580
    screen_w, screen_h = dialog.winfo_screenwidth(), dialog.winfo_screenheight()
    dialog.geometry(f"{width}x{height}"
                    f"+{max(0, (screen_w - width) // 2)}+{max(0, (screen_h - height) // 3)}")
    icon = resource_path(os.path.join("assets", "icon.ico"))
    if os.path.exists(icon):
        try:
            dialog.iconbitmap(default=icon)
        except Exception:
            pass

    tk.Label(dialog, text="免责声明", bg=BG, fg=TEXT,
             font=(FONT_FAMILY, 16, "bold")).pack(anchor="w", padx=22, pady=(16, 0))
    tk.Label(dialog, text="本项目仅供学习与技术交流，不是商业产品。请阅读后输入下面这句话以示知悉。",
             bg=BG, fg=MUTED, font=(FONT_FAMILY, 9)).pack(anchor="w", padx=22, pady=(2, 10))

    box = tk.Frame(dialog, bg=WHITE, highlightthickness=1, highlightbackground=BORDER)
    box.pack(fill="both", expand=True, padx=22)
    body = tk.Text(box, wrap="word", bd=0, padx=14, pady=12, bg=WHITE, fg=TEXT,
                   font=(FONT_FAMILY, 9), height=8, cursor="arrow")
    body.insert("1.0", DISCLAIMER_SUMMARY)
    body.configure(state="disabled")
    body.pack(fill="both", expand=True)

    tk.Label(dialog, text="为确认你已知悉，请完整输入下面这句话：", bg=BG, fg=TEXT,
             font=(FONT_FAMILY, 9, "bold")).pack(anchor="w", padx=22, pady=(12, 0))
    phrase = tk.Label(dialog, text=CONSENT_PHRASE, bg=WHITE, fg=MUTED,
                      font=(FONT_FAMILY, 9), justify="left", wraplength=520,
                      padx=12, pady=8, anchor="w", highlightthickness=1,
                      highlightbackground=BORDER)
    phrase.pack(fill="x", padx=22, pady=(4, 8))

    typed = tk.StringVar()
    entry = tk.Entry(dialog, textvariable=typed, font=(FONT_FAMILY, 10), bd=0,
                     bg=WHITE, fg=TEXT, insertbackground=TEXT, relief="flat",
                     highlightthickness=1, highlightbackground=BORDER,
                     highlightcolor=ACCENT)
    entry.pack(fill="x", padx=22, ipady=7)

    row = tk.Frame(dialog, bg=BG)
    row.pack(fill="x", padx=22, pady=(12, 16))
    result = {"ok": False}

    def refuse():
        result["ok"] = False
        dialog.destroy()

    def agree():
        if not _consent_matches(typed.get()):
            return
        result["ok"] = True
        dialog.destroy()

    def copy_phrase():
        try:
            dialog.clipboard_clear()
            dialog.clipboard_append(CONSENT_PHRASE)
        except Exception:
            pass
        entry.focus_set()

    tk.Button(row, text="退出", command=refuse, font=(FONT_FAMILY, 10), bd=0, relief="flat",
              padx=22, pady=8, bg=WHITE, fg=MUTED, activebackground=WHITE,
              highlightthickness=1, highlightbackground=BORDER,
              cursor="hand2").pack(side="left")
    tk.Button(row, text="复制这句话", command=copy_phrase, font=(FONT_FAMILY, 9), bd=0,
              relief="flat", padx=14, pady=8, bg=WHITE, fg=ACCENT,
              activebackground=WHITE, cursor="hand2").pack(side="left", padx=8)

    agree_button = tk.Button(row, text="同意", font=(FONT_FAMILY, 10, "bold"), bd=0,
                             relief="flat", padx=28, pady=8, bg=BORDER, fg=FAINT,
                             state="disabled", activebackground=ACCENT_DARK)
    agree_button.pack(side="right")

    def refresh(*_args):
        ready = _consent_matches(typed.get())
        agree_button.configure(state="normal" if ready else "disabled",
                               bg=ACCENT if ready else BORDER,
                               fg=WHITE if ready else FAINT,
                               cursor="hand2" if ready else "arrow")
        agree_button.configure(command=agree)

    typed.trace_add("write", refresh)

    dialog.protocol("WM_DELETE_WINDOW", refuse)
    dialog.bind("<Escape>", lambda _event: refuse())
    entry.bind("<Return>", lambda _event: agree())

    dialog.transient(parent)
    dialog.update_idletasks()
    try:
        dialog.grab_set()
    except Exception:
        pass
    entry.focus_set()
    parent.wait_window(dialog)
    return result["ok"]


def main():
    _ensure_writable_temp()
    argv = sys.argv[1:]
    if "--convert" in argv:
        try:
            print(CLI_DISCLAIMER_LINE, file=sys.stderr)
        except Exception:
            pass
        sys.exit(_cli(argv))
    auto_start = "--auto" in argv
    argv = [item for item in argv if item != "--auto"]

    if HAS_DND:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            root = tk.Tk()
    else:
        root = tk.Tk()
    app = App(root)
    if not disclaimer_accepted():
        if not show_disclaimer(root):
            try:
                root.destroy()
            except Exception:
                pass
            return
        app.disclaimer = DISCLAIMER_VERSION
        app._save_settings()
    # files passed on the command line pick the conversion automatically
    accept_now = CONVERSIONS[app.conversion]["accept"]
    initial = [os.path.abspath(a) for a in argv
               if os.path.isfile(a) and Path(a).suffix.lower() in accept_now]
    if not initial:
        for name in CONVERSION_ORDER:
            accept = CONVERSIONS[name]["accept"]
            found = [os.path.abspath(a) for a in argv
                     if os.path.isfile(a) and Path(a).suffix.lower() in accept]
            if found:
                app._select_conversion(name)
                initial = found
                break
    if initial:
        app._add_paths(initial)
    if initial and auto_start:
        root.after(1200, app.start_convert)
    root.mainloop()


if __name__ == "__main__":
    main()
