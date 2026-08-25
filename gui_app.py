# -*- coding: utf-8 -*-
"""
軽量Excelソート GUIアプリケーション
Tkinterを使用した軽快で高速なデスクトップUIを提供します。
"""

import os
import sys
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any, Dict, List, Optional

# 設定マネージャーとソートコア
from config_manager import load_config, save_config, reset_config, get_config_path
from sort_core import process_excel

# ドラッグ&ドロップの安全なインポート
try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False


class SortExcelApp(tk.Tk):
    """ExcelソートツールのメインGUIウィンドウ"""

    def __init__(self):
        super().__init__()
        self.title("Excel受諾確認票 並べ替えツール")
        self.geometry("760x640")
        self.minsize(680, 560)

        # アイコン設定（存在する場合）
        icon_path = Path(__file__).resolve().parent / "icon.ico"
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except Exception:
                pass

        # 設定の読み込み
        self.config: Dict[str, Any] = load_config()

        # 状態変数
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.is_processing = False
        self.last_output_path = ""

        # スタイル設定
        self._setup_styles()

        # UI構築
        self._build_ui()

        # D&D設定
        self._setup_drag_and_drop()

    def _setup_styles(self):
        """UI全体のフォントとテーマスタイルを調整します"""
        self.style = ttk.Style(self)
        
        # Windowsのネイティブテーマ
        available_themes = self.style.theme_names()
        if "vista" in available_themes:
            self.style.theme_use("vista")
        elif "winnative" in available_themes:
            self.style.theme_use("winnative")

        font_family = "Yu Gothic UI" if sys.platform == "win32" else "Helvetica"
        self.default_font = (font_family, 10)
        self.title_font = (font_family, 11, "bold")
        self.btn_font = (font_family, 10, "bold")

        self.option_add("*Font", self.default_font)

        # カスタムスタイル
        self.style.configure("TNotebook.Tab", padding=[16, 8], font=self.default_font)
        self.style.configure("Primary.TButton", font=self.btn_font, padding=[12, 6])
        self.style.configure("Accent.TButton", font=self.btn_font, padding=[16, 8])
        self.style.configure("Header.TLabel", font=self.title_font)

    def _build_ui(self):
        """メインUIコンポーネントを配置します"""
        # タブコントロール
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # タブ1: ソート実行画面
        self.tab_run = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.tab_run, text="  並べ替え実行  ")

        # タブ2: 担当順設定画面
        self.tab_staff = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.tab_staff, text="  担当順の設定  ")

        # タブ3: ソート順 / 詳細設定画面
        self.tab_settings = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.tab_settings, text="  ソート / 書式設定  ")

        # 各タブの中身を初期化
        self._init_tab_run()
        self._init_tab_staff()
        self._init_tab_settings()

        # ステータスバー
        self.status_var = tk.StringVar(value="準備完了")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(6, 2))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ----------------------------------------------------
    # タブ1: ソート実行画面
    # ----------------------------------------------------
    def _init_tab_run(self):
        parent = self.tab_run

        # ファイル選択枠
        file_frame = ttk.LabelFrame(parent, text="入力ファイルの指定", padding=12)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # D&Dガイド領域
        self.drop_area = tk.Label(
            file_frame,
            text="ここにExcelファイルをドラッグ / ドロップしてください\nまたは下の「ファイル選択」ボタンをクリック",
            bg="#f0f4f9",
            fg="#2c3e50",
            relief=tk.GROOVE,
            bd=1,
            padx=10,
            pady=16,
            font=("Yu Gothic UI", 10)
        )
        self.drop_area.pack(fill=tk.X, pady=(0, 10))

        # ファイルパス入力行
        input_row = ttk.Frame(file_frame)
        input_row.pack(fill=tk.X, pady=2)
        ttk.Label(input_row, text="入力パス:").pack(side=tk.LEFT, padx=(0, 6))
        self.entry_input = ttk.Entry(input_row, textvariable=self.input_file_var)
        self.entry_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        btn_browse = ttk.Button(input_row, text="参照...", command=self._browse_input_file)
        btn_browse.pack(side=tk.RIGHT)

        # 出力先設定行
        output_row = ttk.Frame(file_frame)
        output_row.pack(fill=tk.X, pady=(6, 2))
        ttk.Label(output_row, text="出力パス:").pack(side=tk.LEFT, padx=(0, 6))
        self.entry_output = ttk.Entry(output_row, textvariable=self.output_file_var)
        self.entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        btn_browse_out = ttk.Button(output_row, text="変更...", command=self._browse_output_file)
        btn_browse_out.pack(side=tk.RIGHT)

        # 実行ボタン & アクション
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=6)

        self.btn_run = ttk.Button(
            action_frame,
            text="▶ 並べ替えを実行する",
            style="Accent.TButton",
            command=self._start_processing
        )
        self.btn_run.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_open_file = ttk.Button(
            action_frame,
            text="出力ファイルを開く",
            command=self._open_output_file,
            state=tk.DISABLED
        )
        self.btn_open_file.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_open_folder = ttk.Button(
            action_frame,
            text="出力先フォルダを開く",
            command=self._open_output_folder,
            state=tk.DISABLED
        )
        self.btn_open_folder.pack(side=tk.LEFT)

        # プログレスバー
        self.progress_bar = ttk.Progressbar(parent, mode="indeterminate")
        self.progress_bar.pack(fill=tk.X, pady=(8, 4))

        # 実行ログ表示枠
        log_frame = ttk.LabelFrame(parent, text="処理ログ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=8, font=("Consolas", 9), bg="#ffffff")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

    def _setup_drag_and_drop(self):
        """ドラッグ&ドロップをバインドします"""
        if HAS_WINDND:
            try:
                windnd.hook_dropfiles(self.drop_area, func=self._on_drop_files)
                windnd.hook_dropfiles(self, func=self._on_drop_files)
            except Exception as e:
                self._log(f"ドラッグ/ドロップ初期化スキップ: {e}")

    def _on_drop_files(self, files):
        """ドロップされたファイルをセットします"""
        if not files:
            return
        # ファイルパスのデコード
        file_path = files[0]
        if isinstance(file_path, bytes):
            file_path = file_path.decode("mbcs", errors="ignore")
        file_str = str(file_path)
        if file_str.lower().endswith((".xlsx", ".xlsm", ".xls")):
            self._set_input_file(file_str)
        else:
            messagebox.showwarning("ファイル形式エラー", "Excelファイル（.xlsx, .xls, .xlsm）を指定してください。")

    def _browse_input_file(self):
        """入力ファイルの選択ダイアログ"""
        initial_dir = self.config.get("last_input_dir", "")
        if not initial_dir or not Path(initial_dir).exists():
            initial_dir = str(Path.home() / "Desktop")

        path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            initialdir=initial_dir,
            filetypes=[("Excel ファイル", "*.xlsx;*.xlsm;*.xls"), ("すべてのファイル", "*.*")]
        )
        if path:
            self._set_input_file(path)

    def _set_input_file(self, path: str):
        """入力パスとデフォルト出力パスをセットします"""
        self.input_file_var.set(path)
        src = Path(path)
        # 設定の保存ディレクトリを更新
        self.config["last_input_dir"] = str(src.parent)
        save_config(self.config)

        # デフォルト出力先をセット
        default_out = str(src.parent / f"{src.stem}_sorted.xlsx")
        self.output_file_var.set(default_out)
        self._log(f"入力ファイルを設定しました: {src.name}")

    def _browse_output_file(self):
        """出力ファイルの変更ダイアログ"""
        current_out = self.output_file_var.get()
        init_dir = str(Path(current_out).parent) if current_out else str(Path.home() / "Desktop")
        init_file = Path(current_out).name if current_out else "sorted_combined.xlsx"

        path = filedialog.asksaveasfilename(
            title="出力先ファイルの指定",
            initialdir=init_dir,
            initialfile=init_file,
            defaultextension=".xlsx",
            filetypes=[("Excel ファイル", "*.xlsx")]
        )
        if path:
            self.output_file_var.set(path)

    def _log(self, text: str):
        """ログエリアにテキストを追加します（スレッドセーフ）"""
        def append():
            self.log_text.insert(tk.END, text + "\n")
            self.log_text.see(tk.END)
            self.status_var.set(text)
        self.after(0, append)

    def _start_processing(self):
        """ソート処理を別スレッドで開始します"""
        if self.is_processing:
            return

        input_path = self.input_file_var.get().strip()
        if not input_path:
            messagebox.showwarning("入力エラー", "入力Excelファイルを指定してください。")
            return

        if not Path(input_path).exists():
            messagebox.showerror("ファイルエラー", f"指定された入力ファイルが存在しません:\n{input_path}")
            return

        output_path = self.output_file_var.get().strip()
        if not output_path:
            output_path = str(Path(input_path).parent / f"{Path(input_path).stem}_sorted.xlsx")
            self.output_file_var.set(output_path)

        self.is_processing = True
        self.btn_run.config(state=tk.DISABLED)
        self.btn_open_file.config(state=tk.DISABLED)
        self.btn_open_folder.config(state=tk.DISABLED)
        self.progress_bar.start(10)
        self.log_text.delete("1.0", tk.END)
        self._log("処理を開始します...")

        # ワーカースレッドの起動
        thread = threading.Thread(target=self._run_worker, args=(input_path, output_path), daemon=True)
        thread.start()

    def _run_worker(self, input_path: str, output_path: str):
        """バックグラウンドで実行されるソート処理"""
        try:
            out_result = process_excel(
                input_path=input_path,
                output_path=output_path,
                config=self.config,
                log_cb=self._log
            )
            self.last_output_path = out_result
            self.after(0, self._on_process_success, out_result)
        except Exception as e:
            import traceback
            err_trace = traceback.format_exc()
            self._log(f"エラー発生: {e}")
            self.after(0, self._on_process_error, str(e), err_trace)

    def _on_process_success(self, out_path: str):
        """処理完了時のUIコールバック"""
        self.is_processing = False
        self.progress_bar.stop()
        self.btn_run.config(state=tk.NORMAL)
        self.btn_open_file.config(state=tk.NORMAL)
        self.btn_open_folder.config(state=tk.NORMAL)
        self.status_var.set("処理が正常に完了しました")
        messagebox.showinfo("完了", f"並べ替え処理が正常に完了しました。\n\n出力先:\n{out_path}")

    def _on_process_error(self, err_msg: str, trace_msg: str):
        """処理エラー時のUIコールバック"""
        self.is_processing = False
        self.progress_bar.stop()
        self.btn_run.config(state=tk.NORMAL)
        self.status_var.set("エラーが発生しました")
        messagebox.showerror("処理エラー", f"処理中にエラーが発生しました:\n{err_msg}\n\n詳細:\n{trace_msg}")

    def _open_output_file(self):
        """生成されたExcelファイルを既定のアプリで開きます"""
        path = self.output_file_var.get()
        if path and Path(path).exists():
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("起動エラー", f"ファイルを開けませんでした: {e}")
        else:
            messagebox.showwarning("未生成", "出力ファイルが存在しません。")

    def _open_output_folder(self):
        """出力先フォルダをエクスプローラーで開きます"""
        path = self.output_file_var.get()
        folder = Path(path).parent if path else Path.cwd()
        if folder.exists():
            try:
                subprocess.Popen(f'explorer "{folder}"')
            except Exception as e:
                messagebox.showerror("起動エラー", f"フォルダを開けませんでした: {e}")

    # ----------------------------------------------------
    # タブ2: 担当順設定画面
    # ----------------------------------------------------
    def _init_tab_staff(self):
        parent = self.tab_staff

        desc_label = ttk.Label(
            parent,
            text="日付別シートの並べ替えで使用される担当者の優先順位を設定します。\nリストの上にある担当者ほど優先して前に配置されます。"
        )
        desc_label.pack(fill=tk.X, pady=(0, 10))

        content_frame = ttk.Frame(parent)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 左側: 担当者リストボックス
        list_frame = ttk.Frame(content_frame)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self.staff_listbox = tk.Listbox(
            list_frame,
            font=("Yu Gothic UI", 11),
            selectmode=tk.SINGLE,
            activestyle="none",
            exportselection=False
        )
        self.staff_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.staff_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.staff_listbox.config(yscrollcommand=scrollbar.set)

        # 右側: 操作ボタン群
        btn_frame = ttk.Frame(content_frame, width=160)
        btn_frame.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(btn_frame, text="▲▲ 最上部へ", command=lambda: self._move_staff_item("top")).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="▲ 上へ移動", command=lambda: self._move_staff_item("up")).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="▼ 下へ移動", command=lambda: self._move_staff_item("down")).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="▼▼ 最下部へ", command=lambda: self._move_staff_item("bottom")).pack(fill=tk.X, pady=2)

        ttk.Separator(btn_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)

        # 担当者の追加
        ttk.Label(btn_frame, text="担当者名を追加:").pack(anchor=tk.W, pady=(2, 2))
        self.entry_new_staff = ttk.Entry(btn_frame)
        self.entry_new_staff.pack(fill=tk.X, pady=2)
        self.entry_new_staff.bind("<Return>", lambda event: self._add_staff())
        ttk.Button(btn_frame, text="＋ 追加", command=self._add_staff).pack(fill=tk.X, pady=2)

        ttk.Button(btn_frame, text="－ 削除", command=self._delete_staff).pack(fill=tk.X, pady=6)

        ttk.Separator(btn_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)

        ttk.Button(btn_frame, text="五十音順ソート", command=self._sort_staff_alphabetical).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="初期値に戻す", command=self._reset_staff_to_default).pack(fill=tk.X, pady=2)

        # 下部保存ボタン枠
        bottom_frame = ttk.Frame(parent)
        bottom_frame.pack(fill=tk.X, pady=(12, 0))

        btn_save = ttk.Button(bottom_frame, text="✔ 設定を保存する", style="Primary.TButton", command=self._save_all_settings)
        btn_save.pack(side=tk.RIGHT)

        # リスト読み込み
        self._refresh_staff_list()

    def _refresh_staff_list(self):
        """設定からリストボックスへ反映します"""
        self.staff_listbox.delete(0, tk.END)
        for name in self.config.get("staff_order", []):
            self.staff_listbox.insert(tk.END, name)

    def _move_staff_item(self, direction: str):
        """担当者の順番を移動します"""
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
        self.config["staff_order"] = items
        self._refresh_staff_list()
        self.staff_listbox.selection_set(new_idx)
        self.staff_listbox.activate(new_idx)
        self.staff_listbox.see(new_idx)

    def _add_staff(self):
        """新規担当者を追加します"""
        name = self.entry_new_staff.get().strip()
        if not name:
            return
        items = list(self.staff_listbox.get(0, tk.END))
        if name in items:
            messagebox.showinfo("確認", f"「{name}」は既にリストに存在します。")
            return
        items.append(name)
        self.config["staff_order"] = items
        self._refresh_staff_list()
        self.entry_new_staff.delete(0, tk.END)
        # 末尾を選択
        self.staff_listbox.selection_set(len(items) - 1)
        self.staff_listbox.see(len(items) - 1)

    def _delete_staff(self):
        """選択された担当者を削除します"""
        sel = self.staff_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        name = self.staff_listbox.get(idx)
        if messagebox.askyesno("削除確認", f"担当者「{name}」をリストから削除しますか？"):
            items = list(self.staff_listbox.get(0, tk.END))
            items.pop(idx)
            self.config["staff_order"] = items
            self._refresh_staff_list()
            if items:
                new_sel = min(idx, len(items) - 1)
                self.staff_listbox.selection_set(new_sel)

    def _sort_staff_alphabetical(self):
        """五十音順に並べ替えます"""
        items = list(self.staff_listbox.get(0, tk.END))
        items.sort()
        self.config["staff_order"] = items
        self._refresh_staff_list()

    def _reset_staff_to_default(self):
        """担当順を初期設定に戻します"""
        if messagebox.askyesno("確認", "担当順を初期設定に戻しますか？"):
            default = reset_config()
            self.config["staff_order"] = default.get("staff_order", [])
            self._refresh_staff_list()

    # ----------------------------------------------------
    # タブ3: ソート順 / 詳細設定画面
    # ----------------------------------------------------
    def _init_tab_settings(self):
        parent = self.tab_settings

        # 1. 日付シートのソート優先度
        sort_frame = ttk.LabelFrame(parent, text="日付シート内の並べ替え優先順位", padding=12)
        sort_frame.pack(fill=tk.X, pady=(0, 10))

        columns_options = ["識別", "担当", "時間", "お客様氏名", "物件", "場所", "管轄", "依頼内容"]
        order_options = ["昇順", "降順"]

        self.sort_key_vars = []
        sort_keys = self.config.get("sort_keys", [
            {"key": "識別", "ascending": False},
            {"key": "担当", "ascending": True},
            {"key": "時間", "ascending": True}
        ])

        for i in range(3):
            row_frame = ttk.Frame(sort_frame)
            row_frame.pack(fill=tk.X, pady=4)

            ttk.Label(row_frame, text=f"第 {i+1} ソートキー:").pack(side=tk.LEFT, padx=(0, 10))

            key_val = sort_keys[i]["key"] if i < len(sort_keys) else columns_options[0]
            asc_val = "昇順" if (sort_keys[i]["ascending"] if i < len(sort_keys) else True) else "降順"

            var_col = tk.StringVar(value=key_val)
            var_order = tk.StringVar(value=asc_val)

            combo_col = ttk.Combobox(row_frame, textvariable=var_col, values=columns_options, state="readonly", width=14)
            combo_col.pack(side=tk.LEFT, padx=(0, 10))

            combo_order = ttk.Combobox(row_frame, textvariable=var_order, values=order_options, state="readonly", width=8)
            combo_order.pack(side=tk.LEFT)

            self.sort_key_vars.append((var_col, var_order))

        # 2. シート出力オプション
        sheet_frame = ttk.LabelFrame(parent, text="生成するシートの選択", padding=12)
        sheet_frame.pack(fill=tk.X, pady=(0, 10))

        sheet_opts = self.config.get("sheet_options", {})
        self.var_opt_loan = tk.BooleanVar(value=sheet_opts.get("loan_sheet", True))
        self.var_opt_cancel = tk.BooleanVar(value=sheet_opts.get("cancel_sheet", True))
        self.var_opt_combined = tk.BooleanVar(value=sheet_opts.get("combined_sheet", True))
        self.var_opt_date = tk.BooleanVar(value=sheet_opts.get("date_sheets", True))

        ttk.Checkbutton(sheet_frame, text="融資実行日シート（sorted_by_融資実行日）を出力する", variable=self.var_opt_loan).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(sheet_frame, text="金消日・面談日シート（sorted_by_金消日・面談日）を出力する", variable=self.var_opt_cancel).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(sheet_frame, text="統合データシート（統合データ）を出力する", variable=self.var_opt_combined).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(sheet_frame, text="日付ごとの個別シートを自動作成する", variable=self.var_opt_date).pack(anchor=tk.W, pady=2)

        # 3. 書式設定オプション
        format_frame = ttk.LabelFrame(parent, text="表の書式設定", padding=12)
        format_frame.pack(fill=tk.X, pady=(0, 10))

        fmt = self.config.get("formatting", {})
        row_h = fmt.get("row_height", 25)

        h_row = ttk.Frame(format_frame)
        h_row.pack(fill=tk.X, pady=2)
        ttk.Label(h_row, text="行の高さ (標準: 25):").pack(side=tk.LEFT, padx=(0, 10))
        self.var_row_height = tk.IntVar(value=row_h)
        ttk.Spinbox(h_row, from_=15, to=60, textvariable=self.var_row_height, width=6).pack(side=tk.LEFT)

        # 下部ボタン枠
        bottom_frame = ttk.Frame(parent)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))

        btn_reset = ttk.Button(bottom_frame, text="すべての設定を初期値に戻す", command=self._reset_all_settings)
        btn_reset.pack(side=tk.LEFT)

        btn_save = ttk.Button(bottom_frame, text="✔ 設定を保存する", style="Primary.TButton", command=self._save_all_settings)
        btn_save.pack(side=tk.RIGHT)

    def _collect_settings_from_ui(self):
        """UIから設定値を取得してself.configに格納します"""
        # ソートキー
        new_sort_keys = []
        for var_col, var_order in self.sort_key_vars:
            key_name = var_col.get()
            asc = (var_order.get() == "昇順")
            new_sort_keys.append({"key": key_name, "ascending": asc})
        self.config["sort_keys"] = new_sort_keys

        # シート出力設定
        self.config["sheet_options"] = {
            "loan_sheet": self.var_opt_loan.get(),
            "cancel_sheet": self.var_opt_cancel.get(),
            "combined_sheet": self.var_opt_combined.get(),
            "date_sheets": self.var_opt_date.get()
        }

        # 書式設定
        if "formatting" not in self.config:
            self.config["formatting"] = {}
        self.config["formatting"]["row_height"] = self.var_row_height.get()

    def _save_all_settings(self):
        """すべての設定を保存します"""
        self._collect_settings_from_ui()
        ok = save_config(self.config)
        if ok:
            messagebox.showinfo("設定保存", "設定を正常に保存しました。\n次回起動時にもこの設定が適用されます。")
        else:
            messagebox.showerror("設定保存エラー", "設定ファイルの保存に失敗しました。")

    def _reset_all_settings(self):
        """全設定を初期値に戻します"""
        if messagebox.askyesno("設定初期化", "すべての設定を初期状態にリセットしますか？"):
            self.config = reset_config()
            self._refresh_staff_list()
            # UIの変数を更新
            sort_keys = self.config.get("sort_keys", [])
            for i, (var_col, var_order) in enumerate(self.sort_key_vars):
                if i < len(sort_keys):
                    var_col.set(sort_keys[i]["key"])
                    var_order.set("昇順" if sort_keys[i]["ascending"] else "降順")
            sheet_opts = self.config.get("sheet_options", {})
            self.var_opt_loan.set(sheet_opts.get("loan_sheet", True))
            self.var_opt_cancel.set(sheet_opts.get("cancel_sheet", True))
            self.var_opt_combined.set(sheet_opts.get("combined_sheet", True))
            self.var_opt_date.set(sheet_opts.get("date_sheets", True))
            self.var_row_height.set(self.config.get("formatting", {}).get("row_height", 25))
            messagebox.showinfo("リセット完了", "すべての設定を初期値に戻しました。")


def launch_gui(initial_file: Optional[str] = None):
    """GUIアプリケーションを起動します"""
    app = SortExcelApp()
    if initial_file and Path(initial_file).exists():
        app._set_input_file(initial_file)
    app.mainloop()


if __name__ == "__main__":
    initial_input = sys.argv[1] if len(sys.argv) > 1 else None
    launch_gui(initial_input)
