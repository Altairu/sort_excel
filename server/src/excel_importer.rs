use crate::db::Db;
use crate::models::RecordInput;
use anyhow::{Context, Result};
use calamine::{open_workbook_auto, Data, Reader};
use chrono::NaiveDate;
use std::collections::HashMap;
use std::path::Path;

/// Excelセルの値から文字列を取得
fn cell_to_string(data: &Data) -> String {
    match data {
        Data::Empty => String::new(),
        Data::String(s) => s.trim().to_string(),
        Data::Float(f) => {
            // 時間シリアル値 (0.0001 〜 0.9999)
            if *f > 0.0 && *f < 1.0 {
                excel_time_to_string(*f)
            } else if *f >= 30000.0 && *f <= 65000.0 {
                // 日付シリアル値
                excel_date_to_string(*f)
            } else if f.fract() == 0.0 {
                format!("{:.0}", f)
            } else {
                format!("{}", f)
            }
        }
        Data::Int(i) => {
            if *i >= 30000 && *i <= 65000 {
                excel_date_to_string(*i as f64)
            } else {
                format!("{}", i)
            }
        }
        Data::Bool(b) => format!("{}", b),
        Data::DateTime(dt) => {
            let f = dt.as_f64();
            if f > 0.0 && f < 1.0 {
                excel_time_to_string(f)
            } else {
                excel_date_to_string(f)
            }
        }
        Data::DateTimeIso(s) => {
            if s.len() >= 10 {
                s[..10].to_string()
            } else {
                s.clone()
            }
        }
        Data::Error(e) => format!("{:?}", e),
        _ => String::new(),
    }
}

/// Excel時間シリアル値 (0.0〜1.0) を HH:MM に変換
fn excel_time_to_string(serial: f64) -> String {
    let total_sec = (serial * 86400.0).round() as i64;
    let h = (total_sec / 3600) % 24;
    let m = (total_sec % 3600) / 60;
    let s = total_sec % 60;
    if s == 0 {
        format!("{:02}:{:02}", h, m)
    } else {
        format!("{:02}:{:02}:{:02}", h, m, s)
    }
}

/// Excel日付シリアル値 (1900基準) を YYYY-MM-DD に変換
fn excel_date_to_string(serial: f64) -> String {
    let days = serial.floor() as i64;
    let base = NaiveDate::from_ymd_opt(1899, 12, 30);
    if let Some(base_date) = base {
        if let Some(target) = base_date.checked_add_signed(chrono::Duration::days(days)) {
            return target.format("%Y-%m-%d").to_string();
        }
    }
    format!("{}", serial)
}

/// 文字列から日付と補足を分離パース
fn parse_date_and_note(raw: &str) -> (String, String, Option<String>) {
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return (String::new(), String::new(), None);
    }

    if trimmed.contains("リスケ") {
        return (String::new(), trimmed.to_string(), Some("リスケ".into()));
    }
    if trimmed.contains("キャンセル") {
        return (String::new(), trimmed.to_string(), Some("キャンセル".into()));
    }
    if trimmed.contains("見積") {
        return (String::new(), trimmed.to_string(), Some("見積".into()));
    }

    // YYYY-MM-DD または YYYY/MM/DD
    if let Ok(d) = NaiveDate::parse_from_str(trimmed, "%Y-%m-%d") {
        return (d.format("%Y-%m-%d").to_string(), String::new(), None);
    }
    if let Ok(d) = NaiveDate::parse_from_str(trimmed, "%Y/%m/%d") {
        return (d.format("%Y-%m-%d").to_string(), String::new(), None);
    }

    (String::new(), trimmed.to_string(), None)
}

pub fn import_excel_file(path: &Path, db: &Db) -> Result<usize> {
    let mut workbook = open_workbook_auto(path)
        .with_context(|| format!("Excelファイルを開けませんでした: {:?}", path))?;

    let sheet_names = workbook.sheet_names();
    let target_sheet = sheet_names
        .iter()
        .find(|s| *s == "受諾確認票" || *s == "受託確認票" || *s == "Sheet1")
        .cloned()
        .unwrap_or_else(|| sheet_names.first().cloned().unwrap_or_default());

    let range = workbook
        .worksheet_range(&target_sheet)
        .with_context(|| format!("シートが見つかりません: {}", target_sheet))?;

    let mut imported_count = 0;
    let mut col_map: HashMap<String, usize> = HashMap::new();

    for (r_idx, row) in range.rows().enumerate() {
        if r_idx == 0 {
            // ヘッダー行を解析して列名マッピングを作成
            for (c_idx, cell) in row.iter().enumerate() {
                let name = cell_to_string(cell);
                let cleaned = name.trim().replace('　', " ");
                if !cleaned.is_empty() {
                    col_map.insert(cleaned, c_idx);
                }
            }
            tracing::info!("ヘッダー解析結果: {:?}", col_map);
            continue;
        }

        let get = |col_name: &str| -> String {
            if let Some(&idx) = col_map.get(col_name) {
                row.get(idx).map(cell_to_string).unwrap_or_default()
            } else {
                String::new()
            }
        };

        let client_name = get("お客様氏名");
        let property_name = get("物件");
        let form_type = get("形態");
        if client_name.is_empty() && property_name.is_empty() && form_type.is_empty() {
            continue;
        }

        let raw_loan = get("融資実行日");
        let (loan_date, loan_note, status_override) = parse_date_and_note(&raw_loan);

        let raw_settlement = get("金消日・面談日");
        let (settlement_date, settlement_note, _) = parse_date_and_note(&raw_settlement);

        let raw_acceptance = get("受諾日");
        let (acceptance_date, _, _) = parse_date_and_note(&raw_acceptance);

        let mut status = status_override.unwrap_or_else(|| "進行中".to_string());
        if loan_date.is_empty() && settlement_date.is_empty() && status == "進行中" {
            status = "日程未定".to_string();
        }

        let input = RecordInput {
            status: Some(status),
            acceptance_date: Some(acceptance_date),
            loan_date: Some(loan_date),
            loan_date_note: Some(loan_note),
            form_type: Some(form_type),
            client_name: Some(client_name),
            property_name: Some(property_name),
            request_content: Some(get("依頼内容")),
            settlement_date: Some(settlement_date),
            settlement_date_note: Some(settlement_note),
            settlement_place: Some(get("金消場所・面談場所")),
            settlement_time: Some(get("金消時間")),
            staff: Some(get("担当")),
            confirmation: Some(get("意思確認")),
            doc_check: Some(get("申請書類チェック")),
            agency: Some(get("ＡＲＵＨＩ代理店")),
            broker: Some(get("仲介")),
            jurisdiction: Some(get("管轄")),
            opponent_scrivener: Some(get("先方書士")),
            remarks: Some(get("備考")),
            attendance_time: Some(get("立会時間")),
            attendance_place: Some(get("立会場所")),
            attendee: Some(get("立会者")),
            same_day_application: Some(get("当日申請")),
            execution_time: Some(get("実行時間")),
            progress_application: Some(get("申請書")),
            progress_required_docs: Some(get("必要書類")),
            progress_transfer_docs: Some(get("移転書類2")),
            progress_complete_docs: Some(get("完了書類")),
            progress_receive_date: Some(get("受取日")),
            progress_tokyo_complete: Some(get("東京へ完了")),
            progress_payment_date: Some(get("入金日")),
            version: None,
        };

        db.insert_record(&input)?;
        imported_count += 1;
    }

    Ok(imported_count)
}
