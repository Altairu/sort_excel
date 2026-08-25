# -*- coding: utf-8 -*-
"""
設定管理モジュール
担当順やソートキー、書式設定などをJSONファイルとして保存/読み込みします。
"""

import json
from pathlib import Path
from typing import Any, Dict, List

DEFAULT_CONFIG: Dict[str, Any] = {
    "staff_order": [
        "河内", "岩川", "杉田", "正木", "吉田", "中谷", "椙村", "北条", "上野"
    ],
    "sort_keys": [
        {"key": "識別", "ascending": False},
        {"key": "担当", "ascending": True},
        {"key": "時間", "ascending": True}
    ],
    "sheet_options": {
        "loan_sheet": True,
        "cancel_sheet": True,
        "combined_sheet": True,
        "date_sheets": True
    },
    "formatting": {
        "row_height": 25,
        "col_widths": {
            "お客様氏名": 16.43,  # 約115px
            "物件": 22.86,        # 約160px
            "場所": 22.86         # 約160px
        },
        "default_col_width": 11.43  # 約80px
    },
    "last_input_dir": ""
}

CONFIG_FILE_NAME = "sort_excel_config.json"

def get_config_path() -> Path:
    """設定ファイルの保存先パスを返します"""
    return Path(__file__).resolve().parent / CONFIG_FILE_NAME

def load_config() -> Dict[str, Any]:
    """設定ファイルを読み込みます。存在しない場合はデフォルト値を返します"""
    path = get_config_path()
    if not path.exists():
        return dict(DEFAULT_CONFIG)
    
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            merged = dict(DEFAULT_CONFIG)
            for k, v in data.items():
                if isinstance(v, dict) and k in merged and isinstance(merged[k], dict):
                    merged[k] = {**merged[k], **v}
                else:
                    merged[k] = v
            return merged
    except Exception:
        return dict(DEFAULT_CONFIG)

def save_config(config: Dict[str, Any]) -> bool:
    """設定をJSONファイルに保存します"""
    path = get_config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"設定の保存に失敗しました: {e}")
        return False

def reset_config() -> Dict[str, Any]:
    """設定をデフォルト値にリセットして保存します"""
    config = dict(DEFAULT_CONFIG)
    save_config(config)
    return config
