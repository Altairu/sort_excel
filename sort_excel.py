import pandas as pd
import sys


def sort_excel(file_path):
    # Excelファイルの「受諾確認票」シートを読み込み
    df = pd.read_excel(file_path, sheet_name="受諾確認票")

    # 列名の前後の空白を削除
    df.columns = df.columns.str.strip()

    # 日付列の型変換（変換できない場合は NaT となる）
    df['融資実行日'] = pd.to_datetime(df['融資実行日'], errors='coerce')
    df['金消日・面談日'] = pd.to_datetime(df['金消日・面談日'], errors='coerce')

    # ---- (1) 「融資実行日」で並び替えたデータ ----
    sorted_loan = df.sort_values(by='融資実行日')
    # 日付を文字列(YYYY/MM/DD)に変換
    sorted_loan['融資実行日'] = sorted_loan['融資実行日'].dt.strftime('%Y/%m/%d')
    sorted_loan['金消日・面談日'] = sorted_loan['金消日・面談日'].dt.strftime('%Y/%m/%d')
    # 指定された列順に並び替え
    loan_columns = ['融資実行日', '形態', 'お客様氏名', '物件', '依頼内容', '担当', '管轄', '立会時間', '立会場所', '立会者', '当日申請']
    sorted_loan = sorted_loan[loan_columns]

    # ---- (2) 「金消日・面談日」で並び替えたデータ ----
    sorted_cancel = df.sort_values(by='金消日・面談日')
    # 日付を文字列(YYYY/MM/DD)に変換
    sorted_cancel['融資実行日'] = sorted_cancel['融資実行日'].dt.strftime('%Y/%m/%d')
    sorted_cancel['金消日・面談日'] = sorted_cancel['金消日・面談日'].dt.strftime('%Y/%m/%d')
    # 指定された列順に並び替え
    cancel_columns = ['金消日・面談日', '形態', 'お客様氏名', '物件', '依頼内容', '担当', '管轄', '金消時間', '金消場所・面談場所', '意思確認', '融資実行日']
    sorted_cancel = sorted_cancel[cancel_columns]

    # ---- (3) データの統合 ----
    # 融資実行日データに識別列を追加
    sorted_loan['識別'] = '融資実行日'
    sorted_loan = sorted_loan.rename(columns={'融資実行日': '日付', '立会時間': '時間', '立会場所': '場所', '立会者': '確認者', '当日申請': '申請'})

    # 金消日・面談日データに識別列を追加
    sorted_cancel['識別'] = '金消日・面談日'
    sorted_cancel = sorted_cancel.rename(columns={'金消日・面談日': '日付', '金消時間': '時間', '金消場所・面談場所': '場所', '意思確認': '確認者', '融資実行日': '申請'})

    # データを結合
    combined_df = pd.concat([sorted_loan, sorted_cancel], ignore_index=True)

    # 日付で並び替え
    combined_df = combined_df.sort_values(by='日付')

    # 指定された列順に並び替え
    combined_columns = ['日付', '形態', 'お客様氏名', '物件', '依頼内容', '担当', '管轄', '時間', '場所', '確認者', '申請', '識別']
    combined_df = combined_df[combined_columns]

    # Excel出力
    with pd.ExcelWriter('sorted_combined.xlsx') as writer:
        # 既存の3シートを出力
        sorted_loan.to_excel(writer, sheet_name='sorted_by_融資実行日', index=False)
        sorted_cancel.to_excel(writer, sheet_name='sorted_by_金消日・面談日', index=False)
        combined_df.to_excel(writer, sheet_name='統合データ', index=False)

        # ---- (4) 日付ごとに分けたシートを作成 ----
        unique_dates = combined_df['日付'].dropna().unique()  # NaNを除外してユニーク値を取得
        for date in unique_dates:
            date_filtered_df = combined_df[combined_df['日付'] == date]  # 日付でフィルタリング
            sheet_name = date.replace('/', '-')  # シート名に使用できる形式に変換
            date_filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: sort_excel.py <Excelファイルパス>")
        sys.exit(1)

    input_file = sys.argv[1]
    sort_excel(input_file)
