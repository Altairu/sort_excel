# -*- coding: utf-8 -*-
"""
ソート処理コアエンジン
Excelファイルの読み込み、整形、並べ替え、書式設定を行います。
"""

import datetime
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Callable, Dict, Optional, Tuple
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

from config_manager import load_config

# Excelの列見出し（標準化後）
COLS_LOAN = ['融資実行日', '形態', 'お客様氏名', '物件', '依頼内容', '担当', '管轄', '立会時間', '立会場所', '立会者', '当日申請']
COLS_CANCEL = ['金消日・面談日', '形態', 'お客様氏名', '物件', '依頼内容', '担当', '管轄', '金消時間', '金消場所・面談場所', '意思確認', '融資実行日']
COLS_OUT_ORDER = ['日付', '形態', 'お客様氏名', '物件', '依頼内容', '担当', '管轄', '時間', '場所', '確認者', '申請', '識別']

# 全角数字・記号を半角へ揃える変換テーブル
_FULLWIDTH_TRANS = str.maketrans("０１２３４５６７８９／－：", "0123456789/-:")
_RE_YMD = re.compile(r"^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$")
_RE_YMD_JP = re.compile(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$")
_RE_MD = re.compile(r"^(\d{1,2})[/\-](\d{1,2})$")
_RE_MD_JP = re.compile(r"^(\d{1,2})月(\d{1,2})日$")


def _extract_explicit_year(val: Any) -> Optional[int]:
    """値から明示された西暦年（1900〜2100）を取り出します"""
    if val is None:
        return None
    if not isinstance(val, (str, bytes)) and pd.isna(val):
        return None
    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.date)):
        return val.year if 1900 <= val.year <= 2100 else None
    if isinstance(val, str):
        s = val.translate(_FULLWIDTH_TRANS).strip()
        m = _RE_YMD.match(s) or _RE_YMD_JP.match(s)
        if m:
            y = int(m.group(1))
            return y if 1900 <= y <= 2100 else None
    return None


def _infer_reference_year(df: pd.DataFrame) -> int:
    """シート内の日付列から最も多く使われている西暦年を推定します"""
    years = []
    for col in ('融資実行日', '金消日・面談日', '受諾日'):
        if col in df.columns:
            for val in df[col]:
                y = _extract_explicit_year(val)
                if y is not None:
                    years.append(y)
    if years:
        return Counter(years).most_common(1)[0][0]
    return datetime.date.today().year


def _parse_single_date(val: Any, fallback_year: int) -> Any:
    """セル値を安全に日付型へ変換します（年なし M/D の補完と誤解釈防止）"""
    if val is None:
        return pd.NaT
    if not isinstance(val, (str, bytes)) and pd.isna(val):
        return pd.NaT
    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.date)):
        if 1900 <= val.year <= 2100:
            return pd.Timestamp(year=val.year, month=val.month, day=val.day)
        return pd.NaT
    if isinstance(val, (int, float, np.integer, np.floating)):
        num = float(val)
        if 1.0 <= num <= 100000.0:
            try:
                d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(num))
                if 1900 <= d.year <= 2100:
                    return pd.Timestamp(year=d.year, month=d.month, day=d.day)
            except Exception:
                return pd.NaT
        return pd.NaT
    if isinstance(val, str):
        s = val.translate(_FULLWIDTH_TRANS).strip()
        if not s or "\n" in s or "\r" in s:
            return pd.NaT
        m = _RE_YMD.match(s) or _RE_YMD_JP.match(s)
        if m:
            y, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1900 <= y <= 2100:
                try:
                    return pd.Timestamp(year=y, month=month, day=day)
                except ValueError:
                    return pd.NaT
            return pd.NaT
        m_md = _RE_MD.match(s) or _RE_MD_JP.match(s)
        if m_md:
            month, day = int(m_md.group(1)), int(m_md.group(2))
            try:
                return pd.Timestamp(year=fallback_year, month=month, day=day)
            except ValueError:
                return pd.NaT
    return pd.NaT


def _parse_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """DataFrame 内の対象日付列を安全な datetime 型へ変換します"""
    ref_year = _infer_reference_year(df)
    target_cols = [c for c in ('融資実行日', '金消日・面談日') if c in df.columns]
    if not target_cols:
        return df

    # 行ごとに他の日付列から西暦年を参照できるようにします
    year_source_cols = [c for c in ('融資実行日', '金消日・面談日', '受諾日') if c in df.columns]
    row_years = []
    for idx in df.index:
        found_year = None
        for col in year_source_cols:
            y = _extract_explicit_year(df.at[idx, col])
            if y is not None:
                found_year = y
                break
        row_years.append(found_year if found_year is not None else ref_year)

    for col in target_cols:
        parsed = [
            _parse_single_date(val, y)
            for val, y in zip(df[col].tolist(), row_years)
        ]
        df[col] = pd.to_datetime(pd.Series(parsed, index=df.index), errors='coerce')
    return df


def _read_sheet(path: Path, log_cb: Optional[Callable[[str], None]] = None) -> pd.DataFrame:
    """指定されたExcelファイルから対象シートを読み込みます"""
    if log_cb:
        log_cb(f"ファイルを読み込み中: {path.name}")
    excel_file = pd.ExcelFile(path)
    sheets = excel_file.sheet_names

    target_sheet = None
    candidates = ['受諾確認票', '受託確認票', 'Sheet1', 'シート1']
    for candidate in candidates:
        if candidate in sheets:
            target_sheet = candidate
            break

    if target_sheet is None and len(sheets) > 0:
        target_sheet = sheets[0]

    if log_cb:
        log_cb(f"対象シート: {target_sheet}")

    df = pd.read_excel(excel_file, sheet_name=target_sheet)
    df.columns = df.columns.astype(str).str.strip()

    # 日付列のパース（年なし表記の補完と西暦0001年への誤変換防止）
    df = _parse_date_columns(df)
    return df

def _dt_to_str(s: pd.Series) -> pd.Series:
    """日付型Seriesを YYYY/MM/DD 形式の文字列へ変換します"""
    return s.dt.strftime('%Y/%m/%d')

def _parse_time_like(s: pd.Series) -> pd.Series:
    """
    時間文字列（HH:MM、H:MM、HH:MM:SSなど）を秒数に変換します。
    """
    if s is None:
        return pd.Series(dtype='float64')
    ss = s.fillna('')
    def tosec(x):
        t = str(x).strip()
        if not t:
            return np.nan
        parts = t.split(':')
        try:
            h = int(parts[0])
            m = int(parts[1]) if len(parts) > 1 else 0
            sec = int(parts[2]) if len(parts) > 2 else 0
            return h * 3600 + m * 60 + sec
        except Exception:
            return np.nan
    return ss.map(tosec).astype('float64')

def _make_tables(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """融資実行日データ、金消日データ、統合データを生成します"""
    # 融資実行日データ
    loan = df.copy()
    if '融資実行日' in loan.columns:
        loan = loan.sort_values(by='融資実行日')
    loan = loan.assign(識別='融資実行日')
    loan = loan.rename(columns={
        '融資実行日': '日付',
        '立会時間': '時間',
        '立会場所': '場所',
        '立会者': '確認者',
        '当日申請': '申請'
    })
    loan = loan[[c for c in COLS_OUT_ORDER if c in loan.columns]]
    if '日付' in loan.columns:
        loan['日付'] = pd.to_datetime(loan['日付'], errors='coerce')
        loan['日付'] = _dt_to_str(loan['日付'])

    # 金消日・面談日データ
    cancel = df.copy()
    if '金消日・面談日' in cancel.columns:
        cancel = cancel.sort_values(by='金消日・面談日')
    cancel = cancel.assign(識別='金消日・面談日')
    cancel = cancel.rename(columns={
        '金消日・面談日': '日付',
        '金消時間': '時間',
        '金消場所・面談場所': '場所',
        '意思確認': '確認者',
        '融資実行日': '申請'
    })
    cancel = cancel[[c for c in COLS_OUT_ORDER if c in cancel.columns]]
    if '日付' in cancel.columns:
        cancel['日付'] = pd.to_datetime(cancel['日付'], errors='coerce')
        cancel['日付'] = _dt_to_str(cancel['日付'])
    if '申請' in cancel.columns:
        cancel['申請'] = _dt_to_str(pd.to_datetime(cancel['申請'], errors='coerce'))

    # 統合データ
    combined = pd.concat([loan, cancel], ignore_index=True, sort=False)
    if '日付' in combined.columns:
        tmp = pd.to_datetime(combined['日付'], errors='coerce')
        combined = combined.assign(_d=tmp).sort_values(by=['_d']).drop(columns=['_d'])
    combined = combined.reindex(columns=[c for c in COLS_OUT_ORDER if c in combined.columns])

    return loan, cancel, combined

def _sort_for_date_sheet(df_date: pd.DataFrame, config: Dict[str, Any]) -> pd.DataFrame:
    """
    日付シートの並べ替えを設定に基づいて実行します。
    """
    staff_order = config.get("staff_order", [])
    order_dict = {name: i for i, name in enumerate(staff_order)}
    sort_keys_config = config.get("sort_keys", [
        {"key": "識別", "ascending": False},
        {"key": "担当", "ascending": True},
        {"key": "時間", "ascending": True}
    ])

    temp_cols = []
    sort_by_cols = []
    ascending_list = []

    for i, item in enumerate(sort_keys_config):
        col_name = item.get("key", "")
        asc = item.get("ascending", True)
        temp_col = f"_sort_key_{i}"

        if col_name == "担当":
            # 担当順カスタムソート（未掲載は末尾）
            s_val = df_date['担当'].astype(str).map(order_dict).fillna(99999).astype(int)
        elif col_name == "時間":
            # 時間パース
            s_val = _parse_time_like(df_date.get('時間', pd.Series()))
        elif col_name == "識別":
            # 識別（文字列ソート）
            s_val = df_date.get('識別', pd.Series()).astype(str)
        else:
            if col_name in df_date.columns:
                s_val = df_date[col_name].astype(str)
            else:
                s_val = pd.Series([""] * len(df_date), index=df_date.index)

        df_date = df_date.assign(**{temp_col: s_val})
        temp_cols.append(temp_col)
        sort_by_cols.append(temp_col)
        ascending_list.append(asc)

    if sort_by_cols:
        df_sorted = df_date.sort_values(by=sort_by_cols, ascending=ascending_list).drop(columns=temp_cols)
    else:
        df_sorted = df_date

    return df_sorted

def _sanitize_cell_value(val: Any) -> Any:
    """Excelセル用の値整形を行います（欠損値をNoneに、日付・時刻型を文字列に変換）"""
    if val is None:
        return None
    if not isinstance(val, (str, bytes)) and pd.isna(val):
        return None
    if isinstance(val, (pd.Timestamp, datetime.datetime)):
        if val.hour == 0 and val.minute == 0 and val.second == 0:
            return val.strftime('%Y/%m/%d')
        return val.strftime('%Y/%m/%d %H:%M:%S')
    if isinstance(val, datetime.date):
        return val.strftime('%Y/%m/%d')
    if isinstance(val, datetime.time):
        return val.strftime('%H:%M:%S')
    return val


def _apply_excel_styles(out_path: Path, config: Dict[str, Any], log_cb: Optional[Callable[[str], None]] = None):
    """openpyxlを用いて既存Excelファイルへ罫線、列幅、行高、配置スタイルを適用します（後方互換用）"""
    if log_cb:
        log_cb("Excelの書式設定を適用中...")

    formatting = config.get("formatting", {})
    row_height = formatting.get("row_height", 25)
    width_rules = formatting.get("col_widths", {})
    default_width = formatting.get("default_col_width", 11.43)

    wb = load_workbook(out_path)

    dotted = Side(style='dotted')
    border_all = Border(left=dotted, right=dotted, top=dotted, bottom=dotted)
    align = Alignment(horizontal='center', vertical='center', shrink_to_fit=True, wrap_text=False)
    header_font = Font(bold=True)

    for ws in wb.worksheets:
        max_row = ws.max_row
        max_col = ws.max_column

        for r in range(1, max_row + 1):
            ws.row_dimensions[r].height = row_height
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border_all
                if r == 1:
                    cell.font = header_font
                cell.alignment = align

        for c in range(1, max_col + 1):
            col_letter = get_column_letter(c)
            header_val = ws.cell(row=1, column=c).value
            width = width_rules.get(header_val, default_width)
            ws.column_dimensions[col_letter].width = width

    wb.save(out_path)


def process_excel(input_path: str, output_path: Optional[str] = None, config: Optional[Dict[str, Any]] = None, log_cb: Optional[Callable[[str], None]] = None) -> str:
    """
    Excelファイルのソートと書式設定を高速に実行します。
    ディスクI/Oとスタイル適用をシングルパスで行うことで処理時間を短縮します。
    """
    src = Path(input_path)
    if not src.exists():
        raise FileNotFoundError(f"入力ファイルが見つかりません: {src}")

    if config is None:
        config = load_config()

    if output_path is None:
        output_path = str(src.parent / f"{src.stem}_sorted.xlsx")
    
    out_file = Path(output_path)

    # 1. 読み込み
    df = _read_sheet(src, log_cb=log_cb)

    # 2. データ整形
    if log_cb:
        log_cb("データを並べ替え/集計中...")
    loan, cancel, combined = _make_tables(df)

    sheet_opts = config.get("sheet_options", {})
    formatting = config.get("formatting", {})
    row_height = formatting.get("row_height", 25)
    width_rules = formatting.get("col_widths", {})
    default_width = formatting.get("default_col_width", 11.43)

    dotted = Side(style='dotted')
    border_all = Border(left=dotted, right=dotted, top=dotted, bottom=dotted)
    align = Alignment(horizontal='center', vertical='center', shrink_to_fit=True, wrap_text=False)
    header_font = Font(bold=True)

    # 3. 高速シングルパス出力（書き込みと書式設定を同時に実行）
    if log_cb:
        log_cb("シートの作成と書式設定中...")

    wb = Workbook()

    # スタイル定義をキャッシュしてセル単位の重複登録オーバーヘッドを削減
    dummy_ws = wb.active
    c_data = dummy_ws.cell(1, 1)
    c_data.border = border_all
    c_data.alignment = align
    data_style_id = c_data._style

    c_hdr = dummy_ws.cell(1, 2)
    c_hdr.border = border_all
    c_hdr.alignment = align
    c_hdr.font = header_font
    header_style_id = c_hdr._style

    wb.remove(dummy_ws)

    def write_styled_sheet(title: str, target_df: pd.DataFrame) -> None:
        """指定データフレームをシートへ書き込み、スタイルを適用します"""
        ws = wb.create_sheet(title=title)
        cols = list(target_df.columns)

        # 列幅の設定
        for col_idx, col_name in enumerate(cols, 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width_rules.get(col_name, default_width)

        # ヘッダー行の書き込みと書式設定
        ws.append(cols)
        ws.row_dimensions[1].height = row_height
        for cell in ws[1]:
            cell._style = header_style_id

        # データ行の書き込みと書式設定
        if not target_df.empty:
            for r_idx, row_vals in enumerate(target_df.itertuples(index=False, name=None), 2):
                ws.row_dimensions[r_idx].height = row_height
                ws.append([_sanitize_cell_value(v) for v in row_vals])
                for cell in ws[r_idx]:
                    cell._style = data_style_id

    if sheet_opts.get("loan_sheet", True):
        write_styled_sheet('sorted_by_融資実行日', loan)
    if sheet_opts.get("cancel_sheet", True):
        write_styled_sheet('sorted_by_金消日・面談日', cancel)
    if sheet_opts.get("combined_sheet", True):
        write_styled_sheet('統合データ', combined)

    # 日付ごとのシート（ソートキーを一括前計算して高速グループ化）
    if sheet_opts.get("date_sheets", True) and '日付' in combined.columns:
        staff_order = config.get("staff_order", [])
        order_dict = {name: i for i, name in enumerate(staff_order)}
        sort_keys_config = config.get("sort_keys", [
            {"key": "識別", "ascending": False},
            {"key": "担当", "ascending": True},
            {"key": "時間", "ascending": True}
        ])

        comb_work = combined.copy()
        temp_cols = []
        sort_by_cols = []
        ascending_list = []

        for i, item in enumerate(sort_keys_config):
            col_name = item.get("key", "")
            asc = item.get("ascending", True)
            temp_col = f"_sort_key_{i}"

            if col_name == "担当":
                s_val = comb_work['担当'].astype(str).map(order_dict).fillna(99999).astype(int)
            elif col_name == "時間":
                s_val = _parse_time_like(comb_work.get('時間', pd.Series()))
            elif col_name == "識別":
                s_val = comb_work.get('識別', pd.Series()).astype(str)
            else:
                if col_name in comb_work.columns:
                    s_val = comb_work[col_name].astype(str)
                else:
                    s_val = pd.Series([""] * len(comb_work), index=comb_work.index)

            comb_work[temp_col] = s_val
            temp_cols.append(temp_col)
            sort_by_cols.append(temp_col)
            ascending_list.append(asc)

        valid_mask = comb_work['日付'].notna()
        comb_valid = comb_work[valid_mask]
        unique_date_count = comb_valid['日付'].nunique()

        if log_cb:
            log_cb(f"日付別シート作成中 (計 {unique_date_count} 日分)...")

        for date, group in comb_valid.groupby('日付', sort=False):
            if sort_by_cols:
                sorted_group = group.sort_values(by=sort_by_cols, ascending=ascending_list).drop(columns=temp_cols)
            else:
                sorted_group = group.drop(columns=temp_cols)
            sheet_name = str(date).replace('/', '-')[:31]
            write_styled_sheet(sheet_name, sorted_group)

    # 4. ディスクへ保存
    if log_cb:
        log_cb("ファイルを保存中...")
    wb.save(out_file)

    if log_cb:
        log_cb(f"処理が完了しました: {out_file.name}")

    return str(out_file.resolve())
