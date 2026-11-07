# -*- coding: utf-8 -*-
"""
Excel受諾確認票 並べ替えツール GUIアプリケーション
高DPI対応のフラットでシンプルなデスクトップUI、ヘルプ、権利表記、不具合報告機能を提供します。
"""

import ctypes
import os
import platform
import subprocess
import sys
import threading
import urllib.parse
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from config_manager import load_config, save_config, reset_config, get_config_path
import updater
from version import __version__

# ドラッグ&ドロップライブラリの読み込み
try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False


# カラーパレット定義
COLOR_BG_APP = "#F4F6F9"
COLOR_BG_SURFACE = "#FFFFFF"
COLOR_BG_SUBTLE = "#F8FAFC"
COLOR_BORDER = "#DCE1E8"
COLOR_BORDER_FOCUS = "#2563EB"

COLOR_PRIMARY = "#2563EB"
COLOR_PRIMARY_HOVER = "#1D4ED8"
COLOR_PRIMARY_DISABLED = "#94A3B8"
COLOR_PRIMARY_TEXT = "#FFFFFF"

COLOR_SECONDARY_BG = "#F1F5F9"
COLOR_SECONDARY_HOVER = "#E2E8F0"
COLOR_SECONDARY_TEXT = "#1E293B"

COLOR_DANGER_BG = "#FEF2F2"
COLOR_DANGER_HOVER = "#FEE2E2"
COLOR_DANGER_TEXT = "#DC2626"

COLOR_DROP_BG = "#F0F6FF"
COLOR_DROP_HOVER = "#E0EDFF"
COLOR_DROP_BORDER = "#3B82F6"

COLOR_TEXT_MAIN = "#0F172A"
COLOR_TEXT_BODY = "#334155"
COLOR_TEXT_MUTED = "#64748B"

MIT_LICENSE_TEXT = """MIT License

Copyright (c) 2025-2026 -Altair-

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE."""

THIRD_PARTY_NOTICES = """オープンソースソフトウェアの権利表記

本アプリケーションは以下のオープンソースライブラリを利用して構築されています。
各ライブラリの開発者およびコントリビューターに感謝いたします。

* Python / Tkinter
  License: Python Software Foundation License (PSF-2.0)
  Copyright (c) Python Software Foundation. All Rights Reserved.

* pandas
  License: BSD 3-Clause License
  Copyright (c) 2008-2011, AQR Capital Management, LLC, Lambda Foundry, Inc. and PyData Development Team
  Copyright (c) 2011-2026, Open source contributors.

* NumPy
  License: BSD 3-Clause License
  Copyright (c) 2005-2026, NumPy Developers.

* openpyxl
  License: MIT License
  Copyright (c) 2010-2024 openpyxl (Eric Gazoni, Charlie Clark)

* windnd
  License: MIT License
  Copyright (c) 2020 cylin"""


def _enable_high_dpi() -> None:
    """Windows環境での高DPIスケーリングを有効化して画面のにじみを防ぎます"""
    if sys.platform != "win32":
        return
    try:
        # Per-Monitor V2 DPI Awareness を優先適用
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _resolve_resource_path(filename: str) -> Path:
    """スクリプト実行時とPyInstallerのexe実行時の両方でリソースパスを解決します"""
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    candidate = base_dir / filename
    if candidate.exists():
        return candidate
    return Path(__file__).resolve().parent / filename


class FlatButton(tk.Button):
    """ホバーアニメーション付きのフラットデザインボタン"""

    def __init__(
        self,
        master: tk.Widget,
        text: str,
        command: Optional[Callable[[], None]] = None,
        variant: str = "secondary",
        font: Optional[Tuple[Any, ...]] = None,
        padx: int = 14,
        pady: int = 6,
        **kwargs: Any,
    ):
        if variant == "primary":
            self._bg_normal = COLOR_PRIMARY
            self._bg_hover = COLOR_PRIMARY_HOVER
            self._fg_normal = COLOR_PRIMARY_TEXT
            self._fg_disabled = "#E2E8F0"
        elif variant == "danger":
            self._bg_normal = COLOR_DANGER_BG
            self._bg_hover = COLOR_DANGER_HOVER
            self._fg_normal = COLOR_DANGER_TEXT
            self._fg_disabled = COLOR_TEXT_MUTED
        elif variant == "ghost":
            self._bg_normal = COLOR_BG_SURFACE
            self._bg_hover = COLOR_SECONDARY_BG
            self._fg_normal = COLOR_TEXT_BODY
            self._fg_disabled = COLOR_TEXT_MUTED
        else:
            self._bg_normal = COLOR_SECONDARY_BG
            self._bg_hover = COLOR_SECONDARY_HOVER
            self._fg_normal = COLOR_SECONDARY_TEXT
            self._fg_disabled = COLOR_TEXT_MUTED

        self._variant = variant

        super().__init__(
            master,
            text=text,
            command=command,
            bg=self._bg_normal,
            fg=self._fg_normal,
            activebackground=self._bg_hover,
            activeforeground=self._fg_normal,
            disabledforeground=self._fg_disabled,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            font=font,
            padx=padx,
            pady=pady,
            **kwargs,
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _event: tk.Event) -> None:
        if str(self["state"]) != tk.DISABLED:
            self.configure(bg=self._bg_hover)

    def _on_leave(self, _event: tk.Event) -> None:
        if str(self["state"]) != tk.DISABLED:
            self.configure(bg=self._bg_normal)
        elif self._variant == "primary":
            self.configure(bg=COLOR_PRIMARY_DISABLED)

    def set_enabled(self, enabled: bool) -> None:
        """ボタンの有効/無効状態を切り替えます"""
        if enabled:
            self.configure(state=tk.NORMAL, bg=self._bg_normal, cursor="hand2")
        else:
            bg_dis = COLOR_PRIMARY_DISABLED if self._variant == "primary" else COLOR_SECONDARY_BG
            self.configure(state=tk.DISABLED, bg=bg_dis, cursor="arrow")


class SortExcelApp(tk.Tk):
    """Excel受諾確認票 並べ替えツールのメインGUIウィンドウ"""

    def __init__(self):
        _enable_high_dpi()
        super().__init__()

        self.title(f"Excel受諾確認票 並べ替えツール v{__version__}")
        self.configure(bg=COLOR_BG_APP)

        # DPIスケールの算出と設定
        try:
            dpi = float(self.winfo_fpixels("1i"))
        except Exception:
            dpi = 96.0
        self.dpi_scale = max(1.0, min(dpi / 96.0, 2.5))
        try:
            self.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            pass

        win_w = self._px(880)
        win_h = self._px(650)
        self.geometry(f"{win_w}x{win_h}")
        self.minsize(self._px(760), self._px(560))

        # アイコン設定
        icon_path = _resolve_resource_path("icon.ico")
        if icon_path.exists():
            try:
                self.iconbitmap(default=str(icon_path))
            except Exception:
                pass

        # 設定と状態の初期化
        self.config_data: Dict[str, Any] = load_config()
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.status_var = tk.StringVar(value="準備完了")
        self.staff_count_var = tk.StringVar(value="登録数: 0名")
        self.is_processing = False
        self.last_output_path = ""
        self.last_error_summary = ""
        self.last_error_trace = ""
        self.log_history: List[str] = []

        # スタイルとUIの構築
        self._setup_styles()
        self._build_menu()
        self._build_ui()
        self._setup_drag_and_drop()
        self._start_startup_update_check()
        self._start_background_preload()

    def _px(self, value: int) -> int:
        """DPI倍率に応じたピクセル値を返します"""
        return int(round(value * self.dpi_scale))

    def _setup_styles(self) -> None:
        """フォントとttkテーマスタイルを構成します"""
        font_family = "Yu Gothic UI" if sys.platform == "win32" else "Helvetica"
        self.font_family = font_family
        self.font_sm = (font_family, 9)
        self.font_base = (font_family, 10)
        self.font_bold = (font_family, 10, "bold")
        self.font_md = (font_family, 11)
        self.font_card_title = (font_family, 11, "bold")
        self.font_header = (font_family, 13, "bold")
        self.font_drop_title = (font_family, 11, "bold")

        self.option_add("*Font", self.font_base)

        self.style = ttk.Style(self)
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.style.configure(".", background=COLOR_BG_SURFACE, foreground=COLOR_TEXT_MAIN, font=self.font_base)
        self.style.configure("App.TFrame", background=COLOR_BG_APP)
        self.style.configure("Card.TFrame", background=COLOR_BG_SURFACE)

        # チェックボタン
        self.style.configure(
            "Clean.TCheckbutton",
            background=COLOR_BG_SURFACE,
            foreground=COLOR_TEXT_BODY,
            font=self.font_base,
            focuscolor=COLOR_BG_SURFACE,
        )
        self.style.map(
            "Clean.TCheckbutton",
            background=[("active", COLOR_BG_SURFACE)],
            indicatorcolor=[("selected", COLOR_PRIMARY), ("!selected", "#FFFFFF")],
        )

        # コンボボックス
        self.style.configure(
            "Clean.TCombobox",
            fieldbackground="#FFFFFF",
            background=COLOR_SECONDARY_BG,
            foreground=COLOR_TEXT_MAIN,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            arrowcolor=COLOR_TEXT_BODY,
            padding=(self._px(8), self._px(4)),
        )
        self.style.map(
            "Clean.TCombobox",
            fieldbackground=[("readonly", "#FFFFFF")],
            selectbackground=[("readonly", "#FFFFFF")],
            selectforeground=[("readonly", COLOR_TEXT_MAIN)],
            bordercolor=[("focus", COLOR_PRIMARY)],
        )

        # スピンボックス
        self.style.configure(
            "Clean.TSpinbox",
            fieldbackground="#FFFFFF",
            background=COLOR_SECONDARY_BG,
            foreground=COLOR_TEXT_MAIN,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            arrowcolor=COLOR_TEXT_BODY,
            padding=(self._px(8), self._px(4)),
        )

        # プログレスバー
        self.style.configure(
            "Clean.Horizontal.TProgressbar",
            troughcolor="#E2E8F0",
            background=COLOR_PRIMARY,
            bordercolor="#E2E8F0",
            lightcolor=COLOR_PRIMARY,
            darkcolor=COLOR_PRIMARY,
            thickness=self._px(6),
        )

        # スクロールバー
        self.style.configure(
            "Clean.Vertical.TScrollbar",
            gripcount=0,
            background="#CBD5E1",
            troughcolor=COLOR_BG_SUBTLE,
            bordercolor=COLOR_BG_SUBTLE,
            lightcolor="#CBD5E1",
            darkcolor="#CBD5E1",
            arrowcolor=COLOR_TEXT_MUTED,
        )
        self.style.map(
            "Clean.Vertical.TScrollbar",
            background=[("active", "#94A3B8")],
        )

    def _create_card(
        self,
        parent: tk.Widget,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        padding: Tuple[int, int] = (18, 14),
    ) -> Tuple[tk.Frame, tk.Frame]:
        """フラットな白背景のカードコンテナを生成します"""
        outer = tk.Frame(
            parent,
            bg=COLOR_BG_SURFACE,
            highlightbackground=COLOR_BORDER,
            highlightthickness=1,
            bd=0,
        )
        inner = tk.Frame(
            outer,
            bg=COLOR_BG_SURFACE,
            padx=self._px(padding[0]),
            pady=self._px(padding[1]),
        )
        inner.pack(fill=tk.BOTH, expand=True)

        if title:
            lbl_title = tk.Label(
                inner,
                text=title,
                font=self.font_card_title,
                fg=COLOR_TEXT_MAIN,
                bg=COLOR_BG_SURFACE,
                anchor=tk.W,
            )
            lbl_title.pack(fill=tk.X, pady=(0, self._px(2 if subtitle else 6)))

        if subtitle:
            lbl_sub = tk.Label(
                inner,
                text=subtitle,
                font=self.font_sm,
                fg=COLOR_TEXT_MUTED,
                bg=COLOR_BG_SURFACE,
                anchor=tk.W,
                justify=tk.LEFT,
            )
            lbl_sub.pack(fill=tk.X, pady=(0, self._px(8)))

        return outer, inner

    def _create_flat_entry(self, parent: tk.Widget, textvariable: Optional[tk.StringVar] = None) -> Tuple[tk.Frame, tk.Entry]:
        """枠線付きのフラットなテキスト入力欄を生成します"""
        wrapper = tk.Frame(
            parent,
            bg="#FFFFFF",
            highlightbackground="#CBD5E1",
            highlightcolor=COLOR_BORDER_FOCUS,
            highlightthickness=1,
        )
        entry = tk.Entry(
            wrapper,
            textvariable=textvariable,
            font=self.font_base,
            bg="#FFFFFF",
            fg=COLOR_TEXT_MAIN,
            insertbackground=COLOR_TEXT_MAIN,
            relief=tk.FLAT,
            bd=0,
        )
        entry.pack(fill=tk.X, expand=True, padx=self._px(8), pady=self._px(5))
        return wrapper, entry

    # ----------------------------------------------------
    # メインレイアウトとナビゲーション
    # ----------------------------------------------------
    def _build_ui(self) -> None:
        """ヘッダー、ナビゲーションタブ、各ページ、ステータスバーを構築します"""
        # 1. トップヘッダーバー
        header = tk.Frame(self, bg=COLOR_BG_SURFACE, padx=self._px(20), pady=self._px(10))
        header.pack(fill=tk.X)

        title_box = tk.Frame(header, bg=COLOR_BG_SURFACE)
        title_box.pack(side=tk.LEFT)

        tk.Label(
            title_box,
            text="Excel受諾確認票 並べ替えツール",
            font=self.font_header,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_SURFACE,
        ).pack(side=tk.LEFT)

        badge = tk.Label(
            title_box,
            text=f"v{__version__}",
            font=self.font_sm,
            fg=COLOR_PRIMARY,
            bg="#EFF6FF",
            padx=self._px(8),
            pady=self._px(2),
        )
        badge.pack(side=tk.LEFT, padx=(self._px(10), 0))

        header_actions = tk.Frame(header, bg=COLOR_BG_SURFACE)
        header_actions.pack(side=tk.RIGHT)

        FlatButton(
            header_actions,
            text="不具合を報告",
            variant="danger",
            font=self.font_sm,
            padx=self._px(10),
            pady=self._px(4),
            command=self._open_bug_report_dialog,
        ).pack(side=tk.LEFT, padx=(0, self._px(6)))

        FlatButton(
            header_actions,
            text="更新を確認",
            variant="secondary",
            font=self.font_sm,
            padx=self._px(10),
            pady=self._px(4),
            command=self._check_update_manual,
        ).pack(side=tk.LEFT, padx=(0, self._px(6)))

        FlatButton(
            header_actions,
            text="ヘルプ / 権利情報",
            variant="secondary",
            font=self.font_sm,
            padx=self._px(10),
            pady=self._px(4),
            command=lambda: self._switch_tab("help"),
        ).pack(side=tk.LEFT)

        # 2. フラットタブバー
        nav_container = tk.Frame(self, bg=COLOR_BG_SURFACE, padx=self._px(20))
        nav_container.pack(fill=tk.X)

        self.tab_buttons: Dict[str, Tuple[tk.Label, tk.Frame]] = {}
        tab_defs = [
            ("run", "並べ替え実行"),
            ("staff", "担当順の設定"),
            ("settings", "ソート / 書式設定"),
            ("help", "ヘルプ / 権利情報"),
        ]

        for key, label_text in tab_defs:
            tab_item = tk.Frame(nav_container, bg=COLOR_BG_SURFACE, cursor="hand2")
            tab_item.pack(side=tk.LEFT, padx=(0, self._px(8)))

            lbl = tk.Label(
                tab_item,
                text=label_text,
                font=self.font_bold,
                bg=COLOR_BG_SURFACE,
                fg=COLOR_TEXT_MUTED,
                padx=self._px(14),
                pady=self._px(8),
                cursor="hand2",
            )
            lbl.pack(fill=tk.X)

            indicator = tk.Frame(tab_item, bg=COLOR_BG_SURFACE, height=self._px(3))
            indicator.pack(fill=tk.X)

            lbl.bind("<Button-1>", lambda _e, k=key: self._switch_tab(k))
            tab_item.bind("<Button-1>", lambda _e, k=key: self._switch_tab(k))
            self.tab_buttons[key] = (lbl, indicator)

        # ヘッダーとコンテンツの境界線
        tk.Frame(self, bg=COLOR_BORDER, height=1).pack(fill=tk.X)

        # 3. ステータスバー（下部に固定）
        status_outer = tk.Frame(self, bg=COLOR_BG_SURFACE)
        status_outer.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Frame(status_outer, bg=COLOR_BORDER, height=1).pack(fill=tk.X)

        status_inner = tk.Frame(status_outer, bg=COLOR_BG_SURFACE, padx=self._px(16), pady=self._px(6))
        status_inner.pack(fill=tk.X)

        self.status_dot = tk.Label(
            status_inner,
            text="●",
            font=self.font_sm,
            fg="#16A34A",
            bg=COLOR_BG_SURFACE,
        )
        self.status_dot.pack(side=tk.LEFT, padx=(0, self._px(6)))

        tk.Label(
            status_inner,
            textvariable=self.status_var,
            font=self.font_sm,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            status_inner,
            text="MIT License / Copyright (c) 2025-2026 -Altair-",
            font=self.font_sm,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_SURFACE,
        ).pack(side=tk.RIGHT)

        # 4. ページコンテンツ領域
        self.content_area = tk.Frame(self, bg=COLOR_BG_APP, padx=self._px(18), pady=self._px(12))
        self.content_area.pack(fill=tk.BOTH, expand=True)

        self.pages: Dict[str, tk.Frame] = {
            "run": tk.Frame(self.content_area, bg=COLOR_BG_APP),
            "staff": tk.Frame(self.content_area, bg=COLOR_BG_APP),
            "settings": tk.Frame(self.content_area, bg=COLOR_BG_APP),
            "help": tk.Frame(self.content_area, bg=COLOR_BG_APP),
        }

        self._init_tab_run()
        self._init_tab_staff()
        self._init_tab_settings()
        self._init_tab_help()

        self._switch_tab("run")

    def _switch_tab(self, tab_key: str) -> None:
        """表示するタブページを切り替えます"""
        for key, frame in self.pages.items():
            if key == tab_key:
                frame.pack(fill=tk.BOTH, expand=True)
            else:
                frame.pack_forget()

        for key, (lbl, indicator) in self.tab_buttons.items():
            if key == tab_key:
                lbl.configure(fg=COLOR_PRIMARY, font=self.font_bold)
                indicator.configure(bg=COLOR_PRIMARY)
            else:
                lbl.configure(fg=COLOR_TEXT_MUTED, font=self.font_base)
                indicator.configure(bg=COLOR_BG_SURFACE)

    # ----------------------------------------------------
    # タブ1: 並べ替え実行画面
    # ----------------------------------------------------
    def _init_tab_run(self) -> None:
        parent = self.pages["run"]

        # ファイル選択カード
        card_file, file_inner = self._create_card(
            parent,
            title="対象ファイルの指定",
            subtitle="受諾確認票を含むExcelファイルを選択するか、下の枠内へドラッグ & ドロップしてください。",
        )
        card_file.pack(fill=tk.X, pady=(0, self._px(10)))

        # キャンバスによる高解像度ドロップゾーン
        self.drop_canvas = tk.Canvas(
            file_inner,
            height=self._px(76),
            bg=COLOR_DROP_BG,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.drop_canvas.pack(fill=tk.X, pady=(0, self._px(10)))
        self.drop_canvas.bind("<Configure>", self._redraw_drop_canvas)
        self.drop_canvas.bind("<Button-1>", lambda _e: self._browse_input_file())
        self.drop_canvas.bind("<Enter>", lambda _e: self._set_drop_hover(True))
        self.drop_canvas.bind("<Leave>", lambda _e: self._set_drop_hover(False))
        self._drop_hover = False
        self._drop_selected_name = ""

        # 入力ファイル行
        path_grid = tk.Frame(file_inner, bg=COLOR_BG_SURFACE)
        path_grid.pack(fill=tk.X)
        path_grid.columnconfigure(1, weight=1)

        tk.Label(
            path_grid,
            text="入力ファイル",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            width=12,
            anchor=tk.W,
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, self._px(6)))

        in_wrap, self.entry_input = self._create_flat_entry(path_grid, textvariable=self.input_file_var)
        in_wrap.grid(row=0, column=1, sticky=tk.EW, padx=(self._px(6), self._px(8)), pady=(0, self._px(6)))

        FlatButton(
            path_grid,
            text="ファイル選択...",
            variant="secondary",
            font=self.font_base,
            command=self._browse_input_file,
        ).grid(row=0, column=2, sticky=tk.E, pady=(0, self._px(6)))

        # 出力ファイル行
        tk.Label(
            path_grid,
            text="保存先パス",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            width=12,
            anchor=tk.W,
        ).grid(row=1, column=0, sticky=tk.W)

        out_wrap, self.entry_output = self._create_flat_entry(path_grid, textvariable=self.output_file_var)
        out_wrap.grid(row=1, column=1, sticky=tk.EW, padx=(self._px(6), self._px(8)))

        FlatButton(
            path_grid,
            text="保存先変更...",
            variant="secondary",
            font=self.font_base,
            command=self._browse_output_file,
        ).grid(row=1, column=2, sticky=tk.E)

        # 実行アクション行
        action_bar = tk.Frame(file_inner, bg=COLOR_BG_SURFACE)
        action_bar.pack(fill=tk.X, pady=(self._px(12), self._px(4)))

        self.btn_run = FlatButton(
            action_bar,
            text="並べ替えを実行する",
            variant="primary",
            font=self.font_bold,
            padx=self._px(22),
            pady=self._px(7),
            command=self._start_processing,
        )
        self.btn_run.pack(side=tk.LEFT, padx=(0, self._px(10)))

        self.btn_open_file = FlatButton(
            action_bar,
            text="出力ファイルを開く",
            variant="secondary",
            font=self.font_base,
            padx=self._px(14),
            pady=self._px(7),
            command=self._open_output_file,
        )
        self.btn_open_file.pack(side=tk.LEFT, padx=(0, self._px(8)))
        self.btn_open_file.set_enabled(False)

        self.btn_open_folder = FlatButton(
            action_bar,
            text="保存先フォルダを開く",
            variant="secondary",
            font=self.font_base,
            padx=self._px(14),
            pady=self._px(7),
            command=self._open_output_folder,
        )
        self.btn_open_folder.pack(side=tk.LEFT)
        self.btn_open_folder.set_enabled(False)

        # プログレスバー
        self.progress_bar = ttk.Progressbar(
            file_inner,
            mode="indeterminate",
            style="Clean.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill=tk.X, pady=(self._px(6), 0))

        # 処理ログカード
        card_log, log_inner = self._create_card(parent, padding=(18, 12))
        card_log.pack(fill=tk.BOTH, expand=True)

        log_header = tk.Frame(log_inner, bg=COLOR_BG_SURFACE)
        log_header.pack(fill=tk.X, pady=(0, self._px(6)))

        tk.Label(
            log_header,
            text="処理ログ",
            font=self.font_card_title,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_SURFACE,
        ).pack(side=tk.LEFT)

        FlatButton(
            log_header,
            text="ログをクリア",
            variant="secondary",
            font=self.font_sm,
            padx=self._px(10),
            pady=self._px(3),
            command=self._clear_log,
        ).pack(side=tk.RIGHT, padx=(self._px(6), 0))

        FlatButton(
            log_header,
            text="ログをコピー",
            variant="secondary",
            font=self.font_sm,
            padx=self._px(10),
            pady=self._px(3),
            command=self._copy_log_to_clipboard,
        ).pack(side=tk.RIGHT)

        log_box = tk.Frame(
            log_inner,
            bg=COLOR_BG_SUBTLE,
            highlightbackground=COLOR_BORDER,
            highlightthickness=1,
        )
        log_box.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(
            log_box,
            wrap=tk.WORD,
            height=4,
            font=self.font_base,
            bg=COLOR_BG_SUBTLE,
            fg=COLOR_TEXT_BODY,
            relief=tk.FLAT,
            bd=0,
            padx=self._px(10),
            pady=self._px(8),
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        log_scroll = ttk.Scrollbar(
            log_box,
            orient=tk.VERTICAL,
            command=self.log_text.yview,
            style="Clean.Vertical.TScrollbar",
        )
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self._log("対象ファイルを選択して、並べ替えを実行するボタンを押してください。")

    def _set_drop_hover(self, is_hover: bool) -> None:
        self._drop_hover = is_hover
        self.drop_canvas.configure(bg=COLOR_DROP_HOVER if is_hover else COLOR_DROP_BG)
        self._redraw_drop_canvas()

    def _redraw_drop_canvas(self, _event: Optional[tk.Event] = None) -> None:
        """ドロップ領域の破線枠と案内テキストを再描画します"""
        c = self.drop_canvas
        c.delete("all")
        w = c.winfo_width()
        h = c.winfo_height()
        if w <= 10 or h <= 10:
            return

        pad = self._px(6)
        c.create_rectangle(
            pad,
            pad,
            w - pad,
            h - pad,
            outline=COLOR_DROP_BORDER,
            width=max(1, self._px(1)),
            dash=(self._px(6), self._px(4)),
        )

        if self._drop_selected_name:
            main_text = f"選択中のファイル: {self._drop_selected_name}"
            sub_text = "別のファイルをドロップするか、クリックして変更できます"
            main_color = COLOR_PRIMARY
        else:
            main_text = "ここにExcelファイルをドラッグ & ドロップ"
            sub_text = "またはクリックしてファイルを選択 (.xlsx / .xlsm / .xls)"
            main_color = COLOR_TEXT_MAIN

        c.create_text(
            w // 2,
            int(h * 0.38),
            text=main_text,
            font=self.font_drop_title,
            fill=main_color,
        )
        c.create_text(
            w // 2,
            int(h * 0.68),
            text=sub_text,
            font=self.font_sm,
            fill=COLOR_TEXT_MUTED,
        )

    def _setup_drag_and_drop(self) -> None:
        """ドラッグ&ドロップのイベントを登録します"""
        if HAS_WINDND:
            try:
                windnd.hook_dropfiles(self.drop_canvas, func=self._on_drop_files)
                windnd.hook_dropfiles(self, func=self._on_drop_files)
            except Exception as e:
                self._log(f"ドラッグ&ドロップ初期化スキップ: {e}")

    def _on_drop_files(self, files: List[Any]) -> None:
        """ドロップされたファイルを受け取ってセットします"""
        if not files:
            return
        file_path = files[0]
        if isinstance(file_path, bytes):
            file_path = file_path.decode("mbcs", errors="ignore")
        file_str = str(file_path)
        if file_str.lower().endswith((".xlsx", ".xlsm", ".xls")):
            self._switch_tab("run")
            self._set_input_file(file_str)
        else:
            messagebox.showwarning("ファイル形式エラー", "Excelファイル (.xlsx / .xls / .xlsm) を指定してください。")

    def _browse_input_file(self) -> None:
        """入力ファイルの選択ダイアログを表示します"""
        initial_dir = self.config_data.get("last_input_dir", "")
        if not initial_dir or not Path(initial_dir).exists():
            initial_dir = str(Path.home() / "Desktop")

        path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            initialdir=initial_dir,
            filetypes=[("Excel ファイル", "*.xlsx;*.xlsm;*.xls"), ("すべてのファイル", "*.*")],
        )
        if path:
            self._set_input_file(path)

    def _set_input_file(self, path: str) -> None:
        """入力ファイルパスとデフォルトの出力先パスを設定します"""
        self.input_file_var.set(path)
        src = Path(path)
        self._drop_selected_name = src.name
        self._redraw_drop_canvas()

        self.config_data["last_input_dir"] = str(src.parent)
        save_config(self.config_data)

        default_out = str(src.parent / f"{src.stem}_sorted.xlsx")
        self.output_file_var.set(default_out)
        self._log(f"入力ファイルを設定しました: {src.name}")

    def _browse_output_file(self) -> None:
        """出力先ファイルの変更ダイアログを表示します"""
        current_out = self.output_file_var.get()
        init_dir = str(Path(current_out).parent) if current_out else str(Path.home() / "Desktop")
        init_file = Path(current_out).name if current_out else "sorted_combined.xlsx"

        path = filedialog.asksaveasfilename(
            title="保存先ファイルの指定",
            initialdir=init_dir,
            initialfile=init_file,
            defaultextension=".xlsx",
            filetypes=[("Excel ファイル", "*.xlsx")],
        )
        if path:
            self.output_file_var.set(path)

    def _log(self, text: str, status_color: str = "#16A34A") -> None:
        """ログ欄とステータスバーにメッセージを出力します"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {text}"
        self.log_history.append(line)

        def append() -> None:
            self.log_text.insert(tk.END, line + "\n")
            self.log_text.see(tk.END)
            self.status_var.set(text)
            self.status_dot.configure(fg=status_color)

        self.after(0, append)

    def _clear_log(self) -> None:
        """画面上のログを消去します"""
        self.log_text.delete("1.0", tk.END)
        self.log_history.clear()
        self.status_var.set("ログをクリアしました")

    def _copy_log_to_clipboard(self) -> None:
        """処理ログをクリップボードへコピーします"""
        content = self.log_text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showinfo("コピー", "コピーするログがありません。")
            return
        self.clipboard_clear()
        self.clipboard_append(content)
        self.status_var.set("処理ログをクリップボードにコピーしました")

    def _start_processing(self) -> None:
        """ソート処理をバックグラウンドスレッドで開始します"""
        if self.is_processing:
            return

        input_path = self.input_file_var.get().strip()
        if not input_path:
            messagebox.showwarning("入力エラー", "入力するExcelファイルを指定してください。")
            return

        if not Path(input_path).exists():
            messagebox.showerror("ファイルエラー", f"指定された入力ファイルが見つかりません:\n{input_path}")
            return

        output_path = self.output_file_var.get().strip()
        if not output_path:
            output_path = str(Path(input_path).parent / f"{Path(input_path).stem}_sorted.xlsx")
            self.output_file_var.set(output_path)

        # 現在のUI上の設定を反映してから実行
        self._collect_settings_from_ui()

        self.is_processing = True
        self.last_error_summary = ""
        self.last_error_trace = ""
        self.btn_run.set_enabled(False)
        self.btn_open_file.set_enabled(False)
        self.btn_open_folder.set_enabled(False)
        self.progress_bar.start(10)
        self.log_text.delete("1.0", tk.END)
        self._log("並べ替え処理を開始します...", status_color=COLOR_PRIMARY)

        thread = threading.Thread(
            target=self._run_worker,
            args=(input_path, output_path),
            daemon=True,
        )
        thread.start()

    def _run_worker(self, input_path: str, output_path: str) -> None:
        """ワーカースレッドでExcelの整形とソートを実行します"""
        try:
            from sort_core import process_excel
            out_result = process_excel(
                input_path=input_path,
                output_path=output_path,
                config=self.config_data,
                log_cb=lambda msg: self._log(msg, status_color=COLOR_PRIMARY),
            )
            self.last_output_path = out_result
            self.after(0, self._on_process_success, out_result)
        except Exception as e:
            import traceback
            err_trace = traceback.format_exc()
            self._log(f"エラーが発生しました: {e}", status_color=COLOR_DANGER_TEXT)
            self.after(0, self._on_process_error, str(e), err_trace)

    def _on_process_success(self, out_path: str) -> None:
        """処理成功時のUI更新"""
        self.is_processing = False
        self.progress_bar.stop()
        self.btn_run.set_enabled(True)
        self.btn_open_file.set_enabled(True)
        self.btn_open_folder.set_enabled(True)
        self.status_var.set("並べ替え処理が正常に完了しました")
        self.status_dot.configure(fg="#16A34A")
        messagebox.showinfo("処理完了", f"並べ替え処理が正常に完了しました。\n\n保存先:\n{out_path}")

    def _on_process_error(self, err_msg: str, trace_msg: str) -> None:
        """処理失敗時のUI更新とエラー詳細ダイアログの表示"""
        self.is_processing = False
        self.progress_bar.stop()
        self.btn_run.set_enabled(True)
        self.last_error_summary = err_msg
        self.last_error_trace = trace_msg
        self.status_var.set("処理中にエラーが発生しました")
        self.status_dot.configure(fg=COLOR_DANGER_TEXT)
        self._show_error_dialog(err_msg, trace_msg)

    def _show_error_dialog(self, err_msg: str, trace_msg: str) -> None:
        """不具合報告ボタン付きのエラー詳細ダイアログを表示します"""
        dlg = tk.Toplevel(self)
        dlg.title("処理エラー")
        dlg.configure(bg=COLOR_BG_SURFACE)
        dlg.transient(self)
        dlg.geometry(f"{self._px(560)}x{self._px(380)}")
        dlg.minsize(self._px(480), self._px(320))

        body = tk.Frame(dlg, bg=COLOR_BG_SURFACE, padx=self._px(20), pady=self._px(16))
        body.pack(fill=tk.BOTH, expand=True)

        btn_row = tk.Frame(body, bg=COLOR_BG_SURFACE)
        btn_row.pack(side=tk.BOTTOM, fill=tk.X)

        def report_this_error() -> None:
            dlg.destroy()
            self._open_bug_report_dialog(
                initial_title=f"処理エラー: {err_msg[:50]}",
                initial_category="エラー / 動作停止",
                initial_detail="Excelファイルの並べ替え実行時に以下のエラーが発生しました。\n",
                error_trace=trace_msg,
            )

        FlatButton(
            btn_row,
            text="この不具合を報告する",
            variant="danger",
            font=self.font_bold,
            command=report_this_error,
        ).pack(side=tk.LEFT)

        FlatButton(
            btn_row,
            text="閉じる",
            variant="secondary",
            font=self.font_base,
            command=dlg.destroy,
        ).pack(side=tk.RIGHT)

        tk.Label(
            body,
            text="Excelファイルの処理中にエラーが発生しました",
            font=self.font_card_title,
            fg=COLOR_DANGER_TEXT,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X)

        tk.Label(
            body,
            text=err_msg,
            font=self.font_base,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=self._px(500),
        ).pack(fill=tk.X, pady=(self._px(6), self._px(10)))

        tk.Label(
            body,
            text="エラー詳細ログ:",
            font=self.font_sm,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X)

        trace_box = tk.Frame(body, bg=COLOR_BG_SUBTLE, highlightbackground=COLOR_BORDER, highlightthickness=1)
        trace_box.pack(fill=tk.BOTH, expand=True, pady=(self._px(4), self._px(12)))

        txt = tk.Text(
            trace_box,
            wrap=tk.WORD,
            height=6,
            font=self.font_sm,
            bg=COLOR_BG_SUBTLE,
            fg=COLOR_TEXT_BODY,
            relief=tk.FLAT,
            bd=0,
            padx=self._px(8),
            pady=self._px(8),
        )
        txt.insert("1.0", trace_msg)
        txt.configure(state=tk.DISABLED)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(trace_box, orient=tk.VERTICAL, command=txt.yview, style="Clean.Vertical.TScrollbar")
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        txt.configure(yscrollcommand=scroll.set)

        dlg.grab_set()
        dlg.focus_set()

    def _open_output_file(self) -> None:
        """出力されたExcelファイルを既定のアプリケーションで開きます"""
        path = self.output_file_var.get()
        if path and Path(path).exists():
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("起動エラー", f"ファイルを開けませんでした: {e}")
        else:
            messagebox.showwarning("ファイル未生成", "出力ファイルが存在しません。")

    def _open_output_folder(self) -> None:
        """出力先フォルダをエクスプローラーで開きます"""
        path = self.output_file_var.get()
        folder = Path(path).parent if path else Path.cwd()
        if folder.exists():
            try:
                subprocess.Popen(f'explorer "{folder}"')
            except Exception as e:
                messagebox.showerror("起動エラー", f"フォルダを開けませんでした: {e}")

    # ----------------------------------------------------
    # タブ2: 担当順の設定画面
    # ----------------------------------------------------
    def _init_tab_staff(self) -> None:
        parent = self.pages["staff"]

        card_staff, staff_inner = self._create_card(
            parent,
            title="担当者の優先順位設定",
            subtitle="日付別シートの並べ替えで使用する担当者の順序を設定します。リストの上にある担当者ほど優先して上部に配置され、未登録の担当者は末尾に配置されます。",
        )
        card_staff.pack(fill=tk.BOTH, expand=True)

        # 下部保存バー（先に下部に固定）
        bottom_bar = tk.Frame(staff_inner, bg=COLOR_BG_SURFACE)
        bottom_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(self._px(12), 0))

        FlatButton(
            bottom_bar,
            text="設定を保存する",
            variant="primary",
            font=self.font_bold,
            padx=self._px(20),
            pady=self._px(8),
            command=self._save_all_settings,
        ).pack(side=tk.RIGHT)

        body = tk.Frame(staff_inner, bg=COLOR_BG_SURFACE)
        body.pack(fill=tk.BOTH, expand=True)

        # 左カラム: 担当者リスト
        left_col = tk.Frame(body, bg=COLOR_BG_SURFACE)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, self._px(16)))

        list_header = tk.Frame(left_col, bg=COLOR_BG_SURFACE)
        list_header.pack(fill=tk.X, pady=(0, self._px(6)))

        tk.Label(
            list_header,
            text="優先順位リスト (上から優先)",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
        ).pack(side=tk.LEFT)

        tk.Label(
            list_header,
            textvariable=self.staff_count_var,
            font=self.font_sm,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_SURFACE,
        ).pack(side=tk.RIGHT)

        list_box_wrap = tk.Frame(
            left_col,
            bg="#FFFFFF",
            highlightbackground="#CBD5E1",
            highlightthickness=1,
        )
        list_box_wrap.pack(fill=tk.BOTH, expand=True)

        self.staff_listbox = tk.Listbox(
            list_box_wrap,
            font=self.font_md,
            bg="#FFFFFF",
            fg=COLOR_TEXT_MAIN,
            selectbackground="#DBEAFE",
            selectforeground="#1E3A8A",
            selectmode=tk.SINGLE,
            activestyle="none",
            exportselection=False,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            height=8,
        )
        self.staff_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=self._px(6), pady=self._px(6))

        staff_scroll = ttk.Scrollbar(
            list_box_wrap,
            orient=tk.VERTICAL,
            command=self.staff_listbox.yview,
            style="Clean.Vertical.TScrollbar",
        )
        staff_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.staff_listbox.configure(yscrollcommand=staff_scroll.set)

        # 右カラム: 操作パネル
        right_col = tk.Frame(body, bg=COLOR_BG_SURFACE, width=self._px(240))
        right_col.pack(side=tk.RIGHT, fill=tk.Y)
        right_col.pack_propagate(False)

        tk.Label(
            right_col,
            text="順序の移動",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, self._px(6)))

        move_grid = tk.Frame(right_col, bg=COLOR_BG_SURFACE)
        move_grid.pack(fill=tk.X, pady=(0, self._px(10)))
        move_grid.columnconfigure(0, weight=1)
        move_grid.columnconfigure(1, weight=1)

        FlatButton(
            move_grid,
            text="最上部へ",
            variant="secondary",
            font=self.font_base,
            command=lambda: self._move_staff_item("top"),
        ).grid(row=0, column=0, sticky=tk.EW, padx=(0, self._px(4)), pady=(0, self._px(6)))

        FlatButton(
            move_grid,
            text="上へ移動",
            variant="secondary",
            font=self.font_base,
            command=lambda: self._move_staff_item("up"),
        ).grid(row=0, column=1, sticky=tk.EW, padx=(self._px(4), 0), pady=(0, self._px(6)))

        FlatButton(
            move_grid,
            text="下へ移動",
            variant="secondary",
            font=self.font_base,
            command=lambda: self._move_staff_item("down"),
        ).grid(row=1, column=0, sticky=tk.EW, padx=(0, self._px(4)))

        FlatButton(
            move_grid,
            text="最下部へ",
            variant="secondary",
            font=self.font_base,
            command=lambda: self._move_staff_item("bottom"),
        ).grid(row=1, column=1, sticky=tk.EW, padx=(self._px(4), 0))

        tk.Frame(right_col, bg=COLOR_BORDER, height=1).pack(fill=tk.X, pady=self._px(8))

        tk.Label(
            right_col,
            text="担当者の追加と削除",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(self._px(2), self._px(6)))

        add_wrap, self.entry_new_staff = self._create_flat_entry(right_col)
        add_wrap.pack(fill=tk.X, pady=(0, self._px(6)))
        self.entry_new_staff.bind("<Return>", lambda _e: self._add_staff())

        FlatButton(
            right_col,
            text="担当者を追加する",
            variant="secondary",
            font=self.font_bold,
            command=self._add_staff,
        ).pack(fill=tk.X, pady=(0, self._px(6)))

        FlatButton(
            right_col,
            text="選択した担当者を削除",
            variant="danger",
            font=self.font_base,
            command=self._delete_staff,
        ).pack(fill=tk.X, pady=(0, self._px(6)))

        tk.Frame(right_col, bg=COLOR_BORDER, height=1).pack(fill=tk.X, pady=self._px(8))

        FlatButton(
            right_col,
            text="五十音順に並べ替え",
            variant="secondary",
            font=self.font_base,
            command=self._sort_staff_alphabetical,
        ).pack(fill=tk.X, pady=(0, self._px(6)))

        FlatButton(
            right_col,
            text="担当順を初期値に戻す",
            variant="secondary",
            font=self.font_base,
            command=self._reset_staff_to_default,
        ).pack(fill=tk.X)

        self._refresh_staff_list()

    def _refresh_staff_list(self) -> None:
        """設定の担当者リストを画面へ反映します"""
        self.staff_listbox.delete(0, tk.END)
        staff_list = self.config_data.get("staff_order", [])
        for name in staff_list:
            self.staff_listbox.insert(tk.END, name)
        self.staff_count_var.set(f"登録数: {len(staff_list)}名")

    def _move_staff_item(self, direction: str) -> None:
        """選択中の担当者の順序を移動します"""
        sel = self.staff_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        items = list(self.staff_listbox.get(0, tk.END))
        item = items.pop(idx)

        if direction == "up":
            new_idx = max(0, idx - 1)
        elif direction == "down":
            new_idx = min(len(items), idx + 1)
        elif direction == "top":
            new_idx = 0
        elif direction == "bottom":
            new_idx = len(items)
        else:
            new_idx = idx

        items.insert(new_idx, item)
        self.config_data["staff_order"] = items
        self._refresh_staff_list()
        self.staff_listbox.selection_set(new_idx)
        self.staff_listbox.activate(new_idx)
        self.staff_listbox.see(new_idx)

    def _add_staff(self) -> None:
        """新しい担当者を追加します"""
        name = self.entry_new_staff.get().strip()
        if not name:
            return
        items = list(self.staff_listbox.get(0, tk.END))
        if name in items:
            messagebox.showinfo("確認", f"{name} は既にリストに登録されています。")
            return
        items.append(name)
        self.config_data["staff_order"] = items
        self._refresh_staff_list()
        self.entry_new_staff.delete(0, tk.END)
        self.staff_listbox.selection_set(len(items) - 1)
        self.staff_listbox.see(len(items) - 1)

    def _delete_staff(self) -> None:
        """選択された担当者をリストから削除します"""
        sel = self.staff_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        name = self.staff_listbox.get(idx)
        if messagebox.askyesno("削除確認", f"担当者 {name} をリストから削除しますか？"):
            items = list(self.staff_listbox.get(0, tk.END))
            items.pop(idx)
            self.config_data["staff_order"] = items
            self._refresh_staff_list()
            if items:
                new_sel = min(idx, len(items) - 1)
                self.staff_listbox.selection_set(new_sel)

    def _sort_staff_alphabetical(self) -> None:
        """担当者リストを五十音順にソートします"""
        items = list(self.staff_listbox.get(0, tk.END))
        items.sort()
        self.config_data["staff_order"] = items
        self._refresh_staff_list()

    def _reset_staff_to_default(self) -> None:
        """担当者リストを初期設定に戻します"""
        if messagebox.askyesno("確認", "担当順を初期設定に戻しますか？"):
            default = reset_config()
            self.config_data["staff_order"] = default.get("staff_order", [])
            self._refresh_staff_list()

    # ----------------------------------------------------
    # タブ3: ソート / 書式設定画面
    # ----------------------------------------------------
    def _init_tab_settings(self) -> None:
        parent = self.pages["settings"]

        # 下部ボタン行（先に下部に固定）
        bottom_frame = tk.Frame(parent, bg=COLOR_BG_APP)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(self._px(6), 0))

        FlatButton(
            bottom_frame,
            text="すべての設定を初期値に戻す",
            variant="secondary",
            font=self.font_base,
            padx=self._px(14),
            pady=self._px(7),
            command=self._reset_all_settings,
        ).pack(side=tk.LEFT)

        FlatButton(
            bottom_frame,
            text="設定を保存する",
            variant="primary",
            font=self.font_bold,
            padx=self._px(20),
            pady=self._px(7),
            command=self._save_all_settings,
        ).pack(side=tk.RIGHT)

        # 1. 日付別シートのソート優先順位カード
        card_sort, sort_inner = self._create_card(
            parent,
            title="日付別シート内の並べ替え優先順位",
            subtitle="日付ごとのシートを作成する際に、どの列を優先して並べ替えるかを最大3段階まで設定できます。",
            padding=(16, 10),
        )
        card_sort.pack(fill=tk.X, pady=(0, self._px(8)))

        columns_options = ["識別", "担当", "時間", "お客様氏名", "物件", "場所", "管轄", "依頼内容"]
        order_options = ["昇順", "降順"]

        self.sort_key_vars: List[Tuple[tk.StringVar, tk.StringVar]] = []
        sort_keys = self.config_data.get(
            "sort_keys",
            [
                {"key": "識別", "ascending": False},
                {"key": "担当", "ascending": True},
                {"key": "時間", "ascending": True},
            ],
        )

        grid_sort = tk.Frame(sort_inner, bg=COLOR_BG_SURFACE)
        grid_sort.pack(fill=tk.X)

        for i in range(3):
            tk.Label(
                grid_sort,
                text=f"第 {i + 1} 優先キー",
                font=self.font_bold,
                fg=COLOR_TEXT_BODY,
                bg=COLOR_BG_SURFACE,
                width=14,
                anchor=tk.W,
            ).grid(row=i, column=0, sticky=tk.W, pady=self._px(3))

            key_val = sort_keys[i]["key"] if i < len(sort_keys) else columns_options[0]
            asc_val = "昇順" if (sort_keys[i]["ascending"] if i < len(sort_keys) else True) else "降順"

            var_col = tk.StringVar(value=key_val)
            var_order = tk.StringVar(value=asc_val)

            combo_col = ttk.Combobox(
                grid_sort,
                textvariable=var_col,
                values=columns_options,
                state="readonly",
                width=16,
                style="Clean.TCombobox",
            )
            combo_col.grid(row=i, column=1, sticky=tk.W, padx=(self._px(8), self._px(12)), pady=self._px(3))

            combo_order = ttk.Combobox(
                grid_sort,
                textvariable=var_order,
                values=order_options,
                state="readonly",
                width=10,
                style="Clean.TCombobox",
            )
            combo_order.grid(row=i, column=2, sticky=tk.W, pady=self._px(3))

            self.sort_key_vars.append((var_col, var_order))

        # 2. 下段2カラム（左: 出力シート選択カード / 右: 書式とアップデート設定カード）
        lower_row = tk.Frame(parent, bg=COLOR_BG_APP)
        lower_row.pack(fill=tk.BOTH, expand=True)
        lower_row.columnconfigure(0, weight=3)
        lower_row.columnconfigure(1, weight=2)

        card_sheet, sheet_inner = self._create_card(
            lower_row,
            title="出力するシートの選択",
            subtitle="保存されるExcelファイルに含めるシートを選択します。",
            padding=(16, 12),
        )
        card_sheet.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, self._px(8)))

        sheet_opts = self.config_data.get("sheet_options", {})
        self.var_opt_loan = tk.BooleanVar(value=sheet_opts.get("loan_sheet", True))
        self.var_opt_cancel = tk.BooleanVar(value=sheet_opts.get("cancel_sheet", True))
        self.var_opt_combined = tk.BooleanVar(value=sheet_opts.get("combined_sheet", True))
        self.var_opt_date = tk.BooleanVar(value=sheet_opts.get("date_sheets", True))

        ttk.Checkbutton(
            sheet_inner,
            text="融資実行日シート (sorted_by_融資実行日)",
            variable=self.var_opt_loan,
            style="Clean.TCheckbutton",
        ).pack(anchor=tk.W, pady=self._px(2))

        ttk.Checkbutton(
            sheet_inner,
            text="金消日 / 面談日シート (sorted_by_金消日・面談日)",
            variable=self.var_opt_cancel,
            style="Clean.TCheckbutton",
        ).pack(anchor=tk.W, pady=self._px(2))

        ttk.Checkbutton(
            sheet_inner,
            text="統合データシート (統合データ)",
            variable=self.var_opt_combined,
            style="Clean.TCheckbutton",
        ).pack(anchor=tk.W, pady=self._px(2))

        ttk.Checkbutton(
            sheet_inner,
            text="日付ごとの個別シートを自動作成する",
            variable=self.var_opt_date,
            style="Clean.TCheckbutton",
        ).pack(anchor=tk.W, pady=self._px(2))

        card_fmt, fmt_inner = self._create_card(
            lower_row,
            title="書式とアップデート設定",
            subtitle="表の行高と起動時の更新確認を設定します。",
            padding=(16, 12),
        )
        card_fmt.grid(row=0, column=1, sticky=tk.NSEW)

        fmt = self.config_data.get("formatting", {})
        row_h = fmt.get("row_height", 25)

        h_row = tk.Frame(fmt_inner, bg=COLOR_BG_SURFACE)
        h_row.pack(fill=tk.X, pady=(0, self._px(10)))

        tk.Label(
            h_row,
            text="行の高さ (標準: 25)",
            font=self.font_base,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
        ).pack(side=tk.LEFT, padx=(0, self._px(10)))

        self.var_row_height = tk.IntVar(value=row_h)
        ttk.Spinbox(
            h_row,
            from_=15,
            to=60,
            textvariable=self.var_row_height,
            width=8,
            style="Clean.TSpinbox",
        ).pack(side=tk.LEFT)

        update_cfg = self.config_data.get("update", {})
        self.var_auto_update = tk.BooleanVar(value=update_cfg.get("auto_check", True))
        ttk.Checkbutton(
            fmt_inner,
            text="起動時に新しいバージョンを自動確認する",
            variable=self.var_auto_update,
            style="Clean.TCheckbutton",
        ).pack(anchor=tk.W, pady=self._px(2))

    def _collect_settings_from_ui(self) -> None:
        """画面上の設定項目を読み取って config_data に格納します"""
        new_sort_keys = []
        for var_col, var_order in self.sort_key_vars:
            key_name = var_col.get()
            asc = var_order.get() == "昇順"
            new_sort_keys.append({"key": key_name, "ascending": asc})
        self.config_data["sort_keys"] = new_sort_keys

        self.config_data["sheet_options"] = {
            "loan_sheet": self.var_opt_loan.get(),
            "cancel_sheet": self.var_opt_cancel.get(),
            "combined_sheet": self.var_opt_combined.get(),
            "date_sheets": self.var_opt_date.get(),
        }

        if "formatting" not in self.config_data:
            self.config_data["formatting"] = {}
        try:
            self.config_data["formatting"]["row_height"] = int(self.var_row_height.get())
        except Exception:
            self.config_data["formatting"]["row_height"] = 25

        update_cfg = dict(self.config_data.get("update", {}))
        update_cfg["auto_check"] = self.var_auto_update.get()
        self.config_data["update"] = update_cfg

    def _save_all_settings(self) -> None:
        """すべての設定をJSONファイルに保存します"""
        self._collect_settings_from_ui()
        ok = save_config(self.config_data)
        if ok:
            self.status_var.set("設定を保存しました")
            messagebox.showinfo("設定保存", "設定を正常に保存しました。\n次回起動時にもこの設定が適用されます。")
        else:
            messagebox.showerror("保存エラー", "設定ファイルの保存に失敗しました。")

    def _reset_all_settings(self) -> None:
        """全設定を初期状態にリセットします"""
        if messagebox.askyesno("設定初期化", "すべての設定を初期状態にリセットしますか？"):
            self.config_data = reset_config()
            self._refresh_staff_list()
            sort_keys = self.config_data.get("sort_keys", [])
            for i, (var_col, var_order) in enumerate(self.sort_key_vars):
                if i < len(sort_keys):
                    var_col.set(sort_keys[i]["key"])
                    var_order.set("昇順" if sort_keys[i]["ascending"] else "降順")
            sheet_opts = self.config_data.get("sheet_options", {})
            self.var_opt_loan.set(sheet_opts.get("loan_sheet", True))
            self.var_opt_cancel.set(sheet_opts.get("cancel_sheet", True))
            self.var_opt_combined.set(sheet_opts.get("combined_sheet", True))
            self.var_opt_date.set(sheet_opts.get("date_sheets", True))
            self.var_row_height.set(self.config_data.get("formatting", {}).get("row_height", 25))
            self.var_auto_update.set(self.config_data.get("update", {}).get("auto_check", True))
            self.status_var.set("すべての設定を初期値に戻しました")
            messagebox.showinfo("リセット完了", "すべての設定を初期値に戻しました。")

    # ----------------------------------------------------
    # タブ4: ヘルプ / 不具合報告 / 権利情報画面
    # ----------------------------------------------------
    def _init_tab_help(self) -> None:
        parent = self.pages["help"]

        # 上部クイックアクションカード
        card_top, top_inner = self._create_card(
            parent,
            title="サポートとアプリケーション情報",
            subtitle="使い方の確認、不具合の報告、ライセンスや権利表記の確認が行えます。",
            padding=(16, 12),
        )
        card_top.pack(fill=tk.X, pady=(0, self._px(10)))

        btn_row = tk.Frame(top_inner, bg=COLOR_BG_SURFACE)
        btn_row.pack(fill=tk.X)

        FlatButton(
            btn_row,
            text="不具合報告フォームを開く",
            variant="primary",
            font=self.font_bold,
            command=self._open_bug_report_dialog,
        ).pack(side=tk.LEFT, padx=(0, self._px(8)))

        FlatButton(
            btn_row,
            text="GitHub Issues を開く",
            variant="secondary",
            font=self.font_base,
            command=lambda: webbrowser.open(updater.ISSUES_URL),
        ).pack(side=tk.LEFT, padx=(0, self._px(8)))

        FlatButton(
            btn_row,
            text="GitHub リポジトリを開く",
            variant="secondary",
            font=self.font_base,
            command=lambda: webbrowser.open(updater.REPO_URL),
        ).pack(side=tk.LEFT, padx=(0, self._px(8)))

        FlatButton(
            btn_row,
            text="最新バージョンを確認",
            variant="secondary",
            font=self.font_base,
            command=self._check_update_manual,
        ).pack(side=tk.LEFT)

        # 下部ドキュメント切替カード
        card_doc, doc_inner = self._create_card(parent, padding=(16, 12))
        card_doc.pack(fill=tk.BOTH, expand=True)

        sub_nav = tk.Frame(doc_inner, bg=COLOR_BG_SURFACE)
        sub_nav.pack(fill=tk.X, pady=(0, self._px(8)))

        self._help_sections: Dict[str, str] = {
            "guide": self._get_help_guide_text(),
            "trouble": self._get_troubleshooting_text(),
            "rights": self._get_rights_and_license_text(),
        }
        self._help_sub_btns: Dict[str, FlatButton] = {}

        sub_defs = [
            ("guide", "使い方ガイド"),
            ("trouble", "よくある質問 / トラブル対処"),
            ("rights", "権利表記 / ライセンス"),
        ]
        for sec_key, sec_title in sub_defs:
            btn = FlatButton(
                sub_nav,
                text=sec_title,
                variant="secondary",
                font=self.font_bold,
                padx=self._px(14),
                pady=self._px(5),
                command=lambda k=sec_key: self._switch_help_section(k),
            )
            btn.pack(side=tk.LEFT, padx=(0, self._px(8)))
            self._help_sub_btns[sec_key] = btn

        doc_box = tk.Frame(doc_inner, bg=COLOR_BG_SUBTLE, highlightbackground=COLOR_BORDER, highlightthickness=1)
        doc_box.pack(fill=tk.BOTH, expand=True)

        self.help_text_widget = tk.Text(
            doc_box,
            wrap=tk.WORD,
            height=6,
            font=self.font_base,
            bg=COLOR_BG_SUBTLE,
            fg=COLOR_TEXT_BODY,
            relief=tk.FLAT,
            bd=0,
            padx=self._px(14),
            pady=self._px(12),
            spacing1=self._px(2),
            spacing3=self._px(4),
        )
        self.help_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        doc_scroll = ttk.Scrollbar(
            doc_box,
            orient=tk.VERTICAL,
            command=self.help_text_widget.yview,
            style="Clean.Vertical.TScrollbar",
        )
        doc_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.help_text_widget.configure(yscrollcommand=doc_scroll.set)

        self._switch_help_section("guide")

    def _switch_help_section(self, sec_key: str) -> None:
        """ヘルプ画面内の表示セクションを切り替えます"""
        for k, btn in self._help_sub_btns.items():
            if k == sec_key:
                btn.configure(bg=COLOR_PRIMARY, fg=COLOR_PRIMARY_TEXT, activebackground=COLOR_PRIMARY_HOVER)
                btn._bg_normal = COLOR_PRIMARY
                btn._bg_hover = COLOR_PRIMARY_HOVER
                btn._fg_normal = COLOR_PRIMARY_TEXT
            else:
                btn.configure(bg=COLOR_SECONDARY_BG, fg=COLOR_SECONDARY_TEXT, activebackground=COLOR_SECONDARY_HOVER)
                btn._bg_normal = COLOR_SECONDARY_BG
                btn._bg_hover = COLOR_SECONDARY_HOVER
                btn._fg_normal = COLOR_SECONDARY_TEXT

        content = self._help_sections.get(sec_key, "")
        self.help_text_widget.configure(state=tk.NORMAL)
        self.help_text_widget.delete("1.0", tk.END)
        self.help_text_widget.insert("1.0", content)
        self.help_text_widget.configure(state=tk.DISABLED)

    def _get_help_guide_text(self) -> str:
        return (
            "基本的な使い方\n"
            "------------------------------------------------------------\n"
            "* 並べ替え実行タブを開く\n"
            "  画面上部のタブから 並べ替え実行 を選択します。\n\n"
            "* 入力ファイルを指定する\n"
            "  画面中央のドロップ枠にExcelファイルをドラッグ & ドロップするか、枠内をクリックして対象ファイルを選択します。\n"
            "  ファイルを選択すると、同じフォルダ内に [元ファイル名]_sorted.xlsx という保存先パスが自動設定されます。\n\n"
            "* 並べ替えを実行する\n"
            "  並べ替えを実行する ボタンを押すと、データの抽出、統合、日付別シート作成、書式設定が自動で行われます。\n"
            "  完了後は 出力ファイルを開く ボタンや 保存先フォルダを開く ボタンからすぐに結果を確認できます。\n\n\n"
            "対応しているExcelシートと列の仕様\n"
            "------------------------------------------------------------\n"
            "* 自動認識されるシート名\n"
            "  受諾確認票 / 受託確認票 / Sheet1 / シート1 の順に検索し、該当がない場合は先頭のシートを読み込みます。\n\n"
            "* 読み取り対象の主な列\n"
            "  融資実行日 / 金消日・面談日 / 形態 / お客様氏名 / 物件 / 依頼内容 / 担当 / 管轄\n"
            "  立会時間 / 立会場所 / 立会者 / 当日申請 / 金消時間 / 金消場所・面談場所 / 意思確認\n\n"
            "* 出力されるシート構成\n"
            "  * sorted_by_融資実行日 : 融資実行日順のデータ一覧\n"
            "  * sorted_by_金消日・面談日 : 金消日および面談日順のデータ一覧\n"
            "  * 統合データ : 両データを日付昇順で統合した一覧\n"
            "  * 日付別シート : 日付ごとに分割し、指定したソート優先順位と担当順で並べ替えたシート\n\n"
            "* 自動適用される書式\n"
            "  全セルの上下左右中央揃え、縮小して全体を表示、点線罫線、列幅の最適化、指定した行の高さが自動で適用されます。"
        )

    def _get_troubleshooting_text(self) -> str:
        return (
            "よくある質問とトラブルシューティング\n"
            "------------------------------------------------------------\n"
            "* ファイル読み込みや保存時に PermissionError が発生する場合\n"
            "  出力先のExcelファイルが既にExcelで開かれている場合、上書き保存ができません。\n"
            "  Excel側でファイルを閉じてから再度実行するか、保存先パスを別名に変更してください。\n\n"
            "* 期待したシートが読み込まれない場合\n"
            "  入力ファイル内に 受諾確認票 / 受託確認票 / Sheet1 / シート1 のいずれかのシートが含まれているか確認してください。\n\n"
            "* 時間列の並び順が期待通りにならない場合\n"
            "  HH:MM や H:MM:SS 形式の文字列であれば自動で秒数に換算して並べ替えます。\n"
            "  時刻として解釈できない文字列が含まれている行は末尾に配置されます。\n\n"
            "* 自動アップデートが失敗する場合\n"
            "  Program Files などの管理者権限が必要なフォルダに exe を置いている場合、ファイルの差し替えができません。\n"
            "  デスクトップやドキュメントなど、書き込み権限のあるフォルダに exe を配置してください。\n\n"
            "* 問題が解決しない場合や不具合を見つけた場合\n"
            "  画面上部の 不具合を報告 ボタンから、診断情報付きで簡単に報告を作成できます。\n"
            "  GitHubアカウントをお持ちの場合は GitHub Issues で直接送信でき、お持ちでない場合も報告文をクリップボードにコピーして共有できます。"
        )

    def _get_rights_and_license_text(self) -> str:
        return (
            f"アプリケーション権利情報\n"
            f"------------------------------------------------------------\n"
            f"アプリケーション名 : Excel受諾確認票 並べ替えツール (sort_excel)\n"
            f"バージョン         : v{__version__}\n"
            f"著作権表記         : Copyright (c) 2025-2026 -Altair- All Rights Reserved.\n"
            f"ライセンス         : MIT License\n"
            f"配布リポジトリ     : {updater.REPO_URL}\n"
            f"設定ファイル保存先 : {get_config_path()}\n\n\n"
            f"{THIRD_PARTY_NOTICES}\n\n\n"
            f"本ソフトウェアのライセンス全文 (MIT License)\n"
            f"------------------------------------------------------------\n"
            f"{MIT_LICENSE_TEXT}"
        )

    # ----------------------------------------------------
    # 不具合報告ダイアログと診断情報の収集
    # ----------------------------------------------------
    def _build_diagnostic_text(self, error_trace: str = "") -> str:
        """個人情報を除外した安全な環境診断情報と直近ログを生成します"""
        mode_str = "EXE (PyInstaller Frozen)" if updater.is_frozen() else "Python Script"
        os_info = f"{platform.system()} {platform.release()} ({platform.version()})"
        py_info = sys.version.split()[0]

        sort_keys = self.config_data.get("sort_keys", [])
        sort_desc = ", ".join(
            f"{k.get('key')}({'昇順' if k.get('ascending') else '降順'})" for k in sort_keys
        )
        sheet_opts = self.config_data.get("sheet_options", {})

        lines = [
            f"* アプリバージョン: v{__version__}",
            f"* 実行形態: {mode_str}",
            f"* OS環境: {os_info}",
            f"* Python: {py_info}",
            f"* 画面DPI倍率: {int(round(self.dpi_scale * 100))}%",
            f"* ソート設定: {sort_desc}",
            f"* シート出力設定: {sheet_opts}",
        ]

        if self.log_history:
            recent_logs = "\n".join(self.log_history[-15:])
            lines.append("\n直近の処理ログ:")
            lines.append("```text")
            lines.append(recent_logs)
            lines.append("```")

        trace_to_use = error_trace or self.last_error_trace
        if trace_to_use:
            lines.append("\nエラー詳細 (Traceback):")
            lines.append("```text")
            lines.append(trace_to_use.strip())
            lines.append("```")

        return "\n".join(lines)

    def _open_bug_report_dialog(
        self,
        initial_title: str = "",
        initial_category: str = "エラー / 動作停止",
        initial_detail: str = "",
        error_trace: str = "",
    ) -> None:
        """不具合報告を作成してGitHub Issues送信またはコピーを行うダイアログを表示します"""
        dlg = tk.Toplevel(self)
        dlg.title("不具合の報告 / フィードバック")
        dlg.configure(bg=COLOR_BG_SURFACE)
        dlg.transient(self)
        dlg.geometry(f"{self._px(660)}x{self._px(580)}")
        dlg.minsize(self._px(560), self._px(480))

        container = tk.Frame(dlg, bg=COLOR_BG_SURFACE, padx=self._px(20), pady=self._px(16))
        container.pack(fill=tk.BOTH, expand=True)

        diag_text = self._build_diagnostic_text(error_trace=error_trace)

        tk.Label(
            container,
            text="不具合の報告 / フィードバック送信",
            font=self.font_header,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X)

        tk.Label(
            container,
            text="発生した問題の内容を入力してください。GitHub Issues で直接報告を開くか、報告文をコピーして開発者へ共有できます。",
            font=self.font_sm,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
            justify=tk.LEFT,
        ).pack(fill=tk.X, pady=(self._px(2), self._px(10)))

        # 種別と件名
        form_grid = tk.Frame(container, bg=COLOR_BG_SURFACE)
        form_grid.pack(fill=tk.X, pady=(0, self._px(6)))
        form_grid.columnconfigure(1, weight=1)

        tk.Label(
            form_grid,
            text="報告の種別",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            width=12,
            anchor=tk.W,
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, self._px(6)))

        var_category = tk.StringVar(value=initial_category)
        categories = [
            "エラー / 動作停止",
            "並べ替え結果の不備",
            "書式 / 表示の崩れ",
            "使いやすさ / 改善要望",
            "その他",
        ]
        combo_cat = ttk.Combobox(
            form_grid,
            textvariable=var_category,
            values=categories,
            state="readonly",
            width=24,
            style="Clean.TCombobox",
        )
        combo_cat.grid(row=0, column=1, sticky=tk.W, pady=(0, self._px(6)))

        tk.Label(
            form_grid,
            text="概要 (件名)",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            width=12,
            anchor=tk.W,
        ).grid(row=1, column=0, sticky=tk.W)

        var_title = tk.StringVar(value=initial_title)
        title_wrap, _ = self._create_flat_entry(form_grid, textvariable=var_title)
        title_wrap.grid(row=1, column=1, sticky=tk.EW)

        var_attach_diag = tk.BooleanVar(value=True)

        # 下部アクションボタン行（先に下部に固定）
        btn_bar = tk.Frame(container, bg=COLOR_BG_SURFACE)
        btn_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(self._px(10), 0))

        # 詳細入力欄
        tk.Label(
            container,
            text="発生時の状況 / 再現手順",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(self._px(6), self._px(4)))

        detail_wrap = tk.Frame(container, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1)
        detail_wrap.pack(fill=tk.BOTH, expand=True, pady=(0, self._px(6)))

        txt_detail = tk.Text(
            detail_wrap,
            wrap=tk.WORD,
            height=4,
            font=self.font_base,
            bg="#FFFFFF",
            fg=COLOR_TEXT_MAIN,
            relief=tk.FLAT,
            bd=0,
            padx=self._px(8),
            pady=self._px(6),
        )
        default_detail = initial_detail or (
            "発生した問題:\n\n\n"
            "再現手順:\n"
            "* ファイルを読み込む\n"
            "* 並べ替えを実行する\n"
        )
        txt_detail.insert("1.0", default_detail)
        txt_detail.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        detail_scroll = ttk.Scrollbar(
            detail_wrap,
            orient=tk.VERTICAL,
            command=txt_detail.yview,
            style="Clean.Vertical.TScrollbar",
        )
        detail_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        txt_detail.configure(yscrollcommand=detail_scroll.set)

        # 診断情報の添付設定とプレビュー
        ttk.Checkbutton(
            container,
            text="環境診断情報と直近の処理ログを添付する (ファイルパスや担当者名は含まれません)",
            variable=var_attach_diag,
            style="Clean.TCheckbutton",
        ).pack(anchor=tk.W, pady=(0, self._px(4)))

        diag_wrap = tk.Frame(container, bg=COLOR_BG_SUBTLE, highlightbackground=COLOR_BORDER, highlightthickness=1)
        diag_wrap.pack(fill=tk.BOTH, expand=True)

        txt_diag = tk.Text(
            diag_wrap,
            wrap=tk.WORD,
            height=4,
            font=self.font_sm,
            bg=COLOR_BG_SUBTLE,
            fg=COLOR_TEXT_MUTED,
            relief=tk.FLAT,
            bd=0,
            padx=self._px(8),
            pady=self._px(6),
        )
        txt_diag.insert("1.0", diag_text)
        txt_diag.configure(state=tk.DISABLED)
        txt_diag.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        diag_scroll = ttk.Scrollbar(
            diag_wrap,
            orient=tk.VERTICAL,
            command=txt_diag.yview,
            style="Clean.Vertical.TScrollbar",
        )
        diag_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        txt_diag.configure(yscrollcommand=diag_scroll.set)

        def compose_report() -> Tuple[str, str]:
            cat = var_category.get().strip()
            raw_title = var_title.get().strip() or "不具合の報告"
            full_title = f"[{cat}] {raw_title}"
            detail_body = txt_detail.get("1.0", tk.END).strip()

            sections = [
                f"## 種別\n{cat}\n",
                f"## 内容 / 再現手順\n{detail_body}\n",
            ]
            if var_attach_diag.get():
                sections.append(f"## 診断情報\n{diag_text}\n")
            return full_title, "\n".join(sections)

        def submit_github_issue() -> None:
            title_str, body_str = compose_report()
            # クリップボードにも全文を控えとしてコピー
            self.clipboard_clear()
            self.clipboard_append(f"# {title_str}\n\n{body_str}")

            # URL長制限を超えないよう安全に調整
            url_body = body_str
            if len(url_body) > 2800:
                url_body = url_body[:2800] + "\n\n...(全文はクリップボードにコピーされています)"

            params = urllib.parse.urlencode({"title": title_str, "body": url_body})
            webbrowser.open(f"{updater.NEW_ISSUE_URL}?{params}")
            self.status_var.set("GitHub Issues の作成ページを開きました (報告文をクリップボードにもコピー済み)")
            dlg.destroy()

        def copy_report() -> None:
            title_str, body_str = compose_report()
            full_text = f"# {title_str}\n\n{body_str}"
            self.clipboard_clear()
            self.clipboard_append(full_text)
            self.status_var.set("不具合報告文をクリップボードにコピーしました")
            messagebox.showinfo(
                "コピー完了",
                "不具合報告文と診断情報をクリップボードにコピーしました。\nメールやチャット、GitHub Issues に貼り付けて送信できます。",
                parent=dlg,
            )

        def save_report_file() -> None:
            title_str, body_str = compose_report()
            full_text = f"# {title_str}\n\n{body_str}"
            default_name = f"sort_excel_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            save_path = filedialog.asksaveasfilename(
                title="不具合レポートを保存",
                initialdir=str(Path.home() / "Desktop"),
                initialfile=default_name,
                defaultextension=".txt",
                filetypes=[("テキストファイル", "*.txt")],
                parent=dlg,
            )
            if save_path:
                try:
                    Path(save_path).write_text(full_text, encoding="utf-8")
                    messagebox.showinfo("保存完了", f"レポートを保存しました:\n{save_path}", parent=dlg)
                except Exception as e:
                    messagebox.showerror("保存エラー", f"ファイルを保存できませんでした: {e}", parent=dlg)

        FlatButton(
            btn_bar,
            text="GitHub Issues で報告する",
            variant="primary",
            font=self.font_bold,
            padx=self._px(16),
            pady=self._px(7),
            command=submit_github_issue,
        ).pack(side=tk.LEFT, padx=(0, self._px(8)))

        FlatButton(
            btn_bar,
            text="報告文をコピーする",
            variant="secondary",
            font=self.font_bold,
            padx=self._px(14),
            pady=self._px(7),
            command=copy_report,
        ).pack(side=tk.LEFT, padx=(0, self._px(8)))

        FlatButton(
            btn_bar,
            text="ファイルに保存...",
            variant="secondary",
            font=self.font_base,
            padx=self._px(12),
            pady=self._px(7),
            command=save_report_file,
        ).pack(side=tk.LEFT)

        FlatButton(
            btn_bar,
            text="閉じる",
            variant="secondary",
            font=self.font_base,
            padx=self._px(12),
            pady=self._px(7),
            command=dlg.destroy,
        ).pack(side=tk.RIGHT)

        dlg.grab_set()
        dlg.focus_set()

    # ----------------------------------------------------
    # メニューバーとバージョン情報 / 自動アップデート
    # ----------------------------------------------------
    def _build_menu(self) -> None:
        """メニューバーを構築します"""
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Excelファイルを開く...", command=self._browse_input_file)
        file_menu.add_command(label="保存先を変更...", command=self._browse_output_file)
        file_menu.add_separator()
        file_menu.add_command(label="設定を保存", command=self._save_all_settings)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.destroy)
        menubar.add_cascade(label="ファイル", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="使い方ガイド", command=lambda: self._open_help_tab_section("guide"))
        help_menu.add_command(label="よくある質問 / トラブル対処", command=lambda: self._open_help_tab_section("trouble"))
        help_menu.add_separator()
        help_menu.add_command(label="不具合を報告...", command=self._open_bug_report_dialog)
        help_menu.add_command(label="GitHub Issues を開く", command=lambda: webbrowser.open(updater.ISSUES_URL))
        help_menu.add_separator()
        help_menu.add_command(label="更新を確認...", command=self._check_update_manual)
        help_menu.add_separator()
        help_menu.add_command(label="権利表記 / ライセンス", command=lambda: self._open_help_tab_section("rights"))
        help_menu.add_command(label="バージョン情報", command=self._show_about)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)

        self.configure(menu=menubar)

    def _open_help_tab_section(self, sec_key: str) -> None:
        """ヘルプタブに切り替えて指定のセクションを表示します"""
        self._switch_tab("help")
        self._switch_help_section(sec_key)

    def _show_about(self) -> None:
        """バージョン情報と権利表記のダイアログを表示します"""
        dlg = tk.Toplevel(self)
        dlg.title("バージョン情報")
        dlg.configure(bg=COLOR_BG_SURFACE)
        dlg.transient(self)
        dlg.resizable(False, False)
        dlg.geometry(f"{self._px(460)}x{self._px(300)}")

        body = tk.Frame(dlg, bg=COLOR_BG_SURFACE, padx=self._px(24), pady=self._px(20))
        body.pack(fill=tk.BOTH, expand=True)

        btn_row = tk.Frame(body, bg=COLOR_BG_SURFACE)
        btn_row.pack(fill=tk.X, side=tk.BOTTOM)

        FlatButton(
            btn_row,
            text="権利表記 / ライセンス詳細",
            variant="secondary",
            font=self.font_sm,
            command=lambda: (dlg.destroy(), self._open_help_tab_section("rights")),
        ).pack(side=tk.LEFT, padx=(0, self._px(6)))

        FlatButton(
            btn_row,
            text="GitHub リポジトリ",
            variant="secondary",
            font=self.font_sm,
            command=lambda: webbrowser.open(updater.REPO_URL),
        ).pack(side=tk.LEFT)

        FlatButton(
            btn_row,
            text="閉じる",
            variant="primary",
            font=self.font_bold,
            command=dlg.destroy,
        ).pack(side=tk.RIGHT)

        tk.Label(
            body,
            text="Excel受諾確認票 並べ替えツール",
            font=self.font_header,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X)

        tk.Label(
            body,
            text=f"バージョン v{__version__}",
            font=self.font_bold,
            fg=COLOR_PRIMARY,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(self._px(4), self._px(10)))

        tk.Label(
            body,
            text=(
                "Copyright (c) 2025-2026 -Altair- All Rights Reserved.\n"
                "Released under the MIT License.\n\n"
                "受諾確認票のExcelデータを読み込み、用途別の抽出 / 統合 / 日付別ソート / 書式設定を自動で行うデスクトップツールです。"
            ),
            font=self.font_sm,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=self._px(410),
        ).pack(fill=tk.X, pady=(0, self._px(16)))

        dlg.grab_set()
        dlg.focus_set()

    def _start_startup_update_check(self) -> None:
        """起動時にバックグラウンドで最新版を確認します"""
        def worker() -> None:
            updater.cleanup_old_files()
            update_cfg = self.config_data.get("update", {})
            if not update_cfg.get("auto_check", True):
                return
            try:
                info = updater.fetch_latest_release()
            except updater.UpdateError:
                return
            if not updater.is_newer(info):
                return
            if info.tag and info.tag == update_cfg.get("skipped_version", ""):
                return
            self.after(0, self._show_update_dialog, info)

        self.after(1500, lambda: threading.Thread(target=worker, daemon=True).start())

    def _start_background_preload(self) -> None:
        """起動後にバックグラウンドでソート処理ライブラリをプリロードして初回の実行待ちを解消します"""
        def worker() -> None:
            try:
                import sort_core  # noqa: F401
            except Exception:
                pass

        # UI描画が完全に落ち着いた直後に実行
        self.after(300, lambda: threading.Thread(target=worker, daemon=True).start())

    def _check_update_manual(self) -> None:
        """手動で最新バージョンの有無を確認します"""
        self.status_var.set("最新バージョンを確認しています...")

        def worker() -> None:
            try:
                info = updater.fetch_latest_release()
            except updater.UpdateError as e:
                self.after(0, lambda: messagebox.showerror("更新の確認", str(e)))
                self.after(0, lambda: self.status_var.set("準備完了"))
                return
            self.after(0, lambda: self.status_var.set("準備完了"))
            if updater.is_newer(info):
                self.after(0, self._show_update_dialog, info)
            else:
                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "更新の確認",
                        f"お使いのバージョン v{__version__} は最新です。",
                    ),
                )

        threading.Thread(target=worker, daemon=True).start()

    def _show_update_dialog(self, info: "updater.ReleaseInfo") -> None:
        """新しいバージョンの案内ダイアログを表示します"""
        dialog = tk.Toplevel(self)
        dialog.title("アップデートのお知らせ")
        dialog.configure(bg=COLOR_BG_SURFACE)
        dialog.transient(self)
        dialog.geometry(f"{self._px(520)}x{self._px(380)}")
        dialog.minsize(self._px(460), self._px(320))

        frame = tk.Frame(dialog, bg=COLOR_BG_SURFACE, padx=self._px(20), pady=self._px(16))
        frame.pack(fill=tk.BOTH, expand=True)

        btn_row = tk.Frame(frame, bg=COLOR_BG_SURFACE)
        btn_row.pack(side=tk.BOTTOM, fill=tk.X)

        def on_skip() -> None:
            update_cfg = dict(self.config_data.get("update", {}))
            update_cfg["skipped_version"] = info.tag
            self.config_data["update"] = update_cfg
            save_config(self.config_data)
            dialog.destroy()

        def on_open_page() -> None:
            webbrowser.open(info.html_url)
            dialog.destroy()

        can_update = updater.can_self_update(info)
        if can_update:
            FlatButton(
                btn_row,
                text="今すぐ更新する",
                variant="primary",
                font=self.font_bold,
                command=lambda: self._begin_update(info, dialog),
            ).pack(side=tk.RIGHT)
        else:
            FlatButton(
                btn_row,
                text="ダウンロードページを開く",
                variant="primary",
                font=self.font_bold,
                command=on_open_page,
            ).pack(side=tk.RIGHT)

        FlatButton(
            btn_row,
            text="後で",
            variant="secondary",
            font=self.font_base,
            command=dialog.destroy,
        ).pack(side=tk.RIGHT, padx=self._px(6))

        FlatButton(
            btn_row,
            text="このバージョンをスキップ",
            variant="secondary",
            font=self.font_sm,
            command=on_skip,
        ).pack(side=tk.LEFT)

        tk.Label(
            frame,
            text=f"新しいバージョン {info.tag} が公開されています",
            font=self.font_card_title,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X)

        tk.Label(
            frame,
            text=f"現在のバージョン: v{__version__}",
            font=self.font_sm,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(self._px(2), self._px(10)))

        tk.Label(
            frame,
            text="更新内容:",
            font=self.font_bold,
            fg=COLOR_TEXT_BODY,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X)

        notes_frame = tk.Frame(frame, bg=COLOR_BG_SUBTLE, highlightbackground=COLOR_BORDER, highlightthickness=1)
        notes_frame.pack(fill=tk.BOTH, expand=True, pady=(self._px(4), self._px(10)))

        notes = tk.Text(
            notes_frame,
            wrap=tk.WORD,
            height=6,
            font=self.font_sm,
            bg=COLOR_BG_SUBTLE,
            fg=COLOR_TEXT_BODY,
            relief=tk.FLAT,
            bd=0,
            padx=self._px(8),
            pady=self._px(6),
        )
        notes.insert("1.0", info.notes or "変更内容の記載はありません。")
        notes.configure(state=tk.DISABLED)
        notes.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        notes_scroll = ttk.Scrollbar(
            notes_frame,
            orient=tk.VERTICAL,
            command=notes.yview,
            style="Clean.Vertical.TScrollbar",
        )
        notes_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        notes.configure(yscrollcommand=notes_scroll.set)

        if not can_update:
            reason = (
                "Python スクリプトとして実行中のため、自動更新は行いません。"
                if not updater.is_frozen()
                else "このリリースは自動更新に対応していません。"
            )
            tk.Label(
                frame,
                text=reason + "\nダウンロードページから最新版を入手してください。",
                font=self.font_sm,
                fg=COLOR_TEXT_MUTED,
                bg=COLOR_BG_SURFACE,
                anchor=tk.W,
                justify=tk.LEFT,
            ).pack(fill=tk.X, pady=(0, self._px(10)))

        dialog.grab_set()
        dialog.focus_set()

    def _begin_update(self, info: "updater.ReleaseInfo", dialog: tk.Toplevel) -> None:
        """最新版のダウンロードと差し替えを実行します"""
        if self.is_processing:
            messagebox.showwarning(
                "更新",
                "並べ替え処理の実行中です。処理が終わってから更新してください。",
                parent=dialog,
            )
            return
        dialog.destroy()

        win = tk.Toplevel(self)
        win.title("アップデート")
        win.configure(bg=COLOR_BG_SURFACE)
        win.transient(self)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = tk.Frame(win, bg=COLOR_BG_SURFACE, padx=self._px(20), pady=self._px(18))
        frame.pack(fill=tk.BOTH, expand=True)

        status = tk.StringVar(value=f"{info.tag} をダウンロードしています...")
        tk.Label(
            frame,
            textvariable=status,
            font=self.font_base,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_SURFACE,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, self._px(10)))

        bar = ttk.Progressbar(
            frame,
            mode="determinate",
            length=self._px(380),
            maximum=100,
            style="Clean.Horizontal.TProgressbar",
        )
        bar.pack(fill=tk.X)
        win.grab_set()

        def on_progress(done: int, total: int) -> None:
            def ui() -> None:
                if total > 0:
                    bar["value"] = done * 100 / total
                    status.set(
                        f"{info.tag} をダウンロードしています... {done // (1024 * 1024)} / {total // (1024 * 1024)} MB"
                    )

            self.after(0, ui)

        def worker() -> None:
            try:
                new_path = updater.download_update(info, progress_cb=on_progress)
            except updater.UpdateError as e:
                self.after(0, self._on_update_failed, win, str(e))
                return
            self.after(0, self._finish_update, win, new_path)

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_failed(self, win: tk.Toplevel, message: str) -> None:
        """更新失敗時のエラーハンドリング"""
        win.grab_release()
        win.destroy()
        messagebox.showerror("アップデート失敗", f"{message}\n\n現在のバージョンのまま使用できます。")

    def _finish_update(self, win: tk.Toplevel, new_path: Path) -> None:
        """exeを差し替えて再起動します"""
        try:
            updater.apply_update(new_path)
        except updater.UpdateError as e:
            self._on_update_failed(win, str(e))
            return
        win.grab_release()
        win.destroy()
        self.destroy()


def launch_gui(initial_file: Optional[str] = None) -> None:
    """GUIアプリケーションを起動します"""
    app = SortExcelApp()
    if initial_file and Path(initial_file).exists():
        app._set_input_file(initial_file)
    app.mainloop()


if __name__ == "__main__":
    initial_input = sys.argv[1] if len(sys.argv) > 1 else None
    launch_gui(initial_input)
