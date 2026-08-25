#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sort_excel.py
- Excelの「受諾確認票」シートを読み込み、用途別に並べ替え・書式設定を行って出力
- GUIモード（引数なし起動）とCLI / ドロップ実行の両方に対応
- 担当順やソート順の設定変更・保存に対応
"""

import sys
import ctypes
from pathlib import Path
from typing import Optional

# 設定マネージャーとコアロジック
from config_manager import load_config
from sort_core import process_excel


def _show_msg(title: str, text: str, style: int = 0):
    """Windowsネイティブのメッセージボックスを表示します"""
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, style)
    except Exception:
        print(f"[{title}] {text}")


def sort_excel(input_path: str, output_path: Optional[str] = None):
    """
    既存の外部呼び出し互換関数。
    Excelファイルを読み込み、並べ替えと書式設定を行って保存します。
    """
    if output_path is None:
        output_path = 'sorted_combined.xlsx'
    return process_excel(input_path=input_path, output_path=output_path)


def main():
    """エントリーポイント"""
    if len(sys.argv) < 2:
        # 引数なしの場合はGUIを起動
        from gui_app import launch_gui
        launch_gui()
    else:
        # 引数がある場合（ファイルが渡された場合）
        arg_path = sys.argv[1]
        
        # もし --gui フラグまたはGUIで開く挙動を希望する場合の処理
        if arg_path in ("--gui", "-g"):
            from gui_app import launch_gui
            target_file = sys.argv[2] if len(sys.argv) > 2 else None
            launch_gui(target_file)
            return

        # ファイルが存在する場合はGUIに読み込ませて起動
        from gui_app import launch_gui
        launch_gui(arg_path)


if __name__ == '__main__':
    main()
