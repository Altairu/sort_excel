use crate::models::{AppConfig, Record};
use anyhow::Result;
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Workbook};
use serde::Serialize;
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::thread;
use std::time::Duration;

/// 日付別シートやWeb画面のソート済みビューで共有する行データ構造
#[derive(Debug, Clone, Serialize)]
pub struct ExportRow {
    pub record_id: i64,
    pub status: String,
    pub date: String,
    pub form_type: String,
    pub client_name: String,
    pub property_name: String,
    pub request_content: String,
    pub staff: String,
    pub jurisdiction: String,
    pub time: String,
    pub place: String,
    pub confirmer: String,
    pub application: String,
    pub category: String, // 融資実行日 or 金消日・面談日
}

/// Web画面表示およびExcel出力で共通利用するソート済みデータ一式
#[derive(Debug, Clone, Serialize)]
pub struct SortedViewsData {
    pub loan_rows: Vec<ExportRow>,
    pub cancel_rows: Vec<ExportRow>,
    pub combined_rows: Vec<ExportRow>,
    pub date_groups: BTreeMap<String, Vec<ExportRow>>,
}

/// 時間文字列 (HH:MM や HH:MM:SS) を秒数に変換して比較可能にする
fn parse_time_to_seconds(t: &str) -> i64 {
    let t = t.trim();
    if t.is_empty() || t == "-" {
        return 999999;
    }
    let parts: Vec<&str> = t.split(':').collect();
    if parts.is_empty() {
        return 999999;
    }
    let h: i64 = parts[0].parse().unwrap_or(99);
    let m: i64 = parts.get(1).and_then(|s| s.parse().ok()).unwrap_or(0);
    let s: i64 = parts.get(2).and_then(|s| s.parse().ok()).unwrap_or(0);
    h * 3600 + m * 60 + s
}

/// レコード一覧と設定からソート済みデータ一式を構築
pub fn build_sorted_views(records: &[Record], config: &AppConfig) -> SortedViewsData {
    let mut loan_rows: Vec<ExportRow> = Vec::new();
    let mut cancel_rows: Vec<ExportRow> = Vec::new();

    for rec in records {
        if !rec.loan_date.is_empty() {
            loan_rows.push(ExportRow {
                record_id: rec.id,
                status: rec.status.clone(),
                date: rec.loan_date.clone(),
                form_type: rec.form_type.clone(),
                client_name: rec.client_name.clone(),
                property_name: rec.property_name.clone(),
                request_content: rec.request_content.clone(),
                staff: rec.staff.clone(),
                jurisdiction: rec.jurisdiction.clone(),
                time: rec.attendance_time.clone(),
                place: rec.attendance_place.clone(),
                confirmer: rec.attendee.clone(),
                application: rec.same_day_application.clone(),
                category: "融資実行日".to_string(),
            });
        }
        if !rec.settlement_date.is_empty() {
            cancel_rows.push(ExportRow {
                record_id: rec.id,
                status: rec.status.clone(),
                date: rec.settlement_date.clone(),
                form_type: rec.form_type.clone(),
                client_name: rec.client_name.clone(),
                property_name: rec.property_name.clone(),
                request_content: rec.request_content.clone(),
                staff: rec.staff.clone(),
                jurisdiction: rec.jurisdiction.clone(),
                time: rec.settlement_time.clone(),
                place: rec.settlement_place.clone(),
                confirmer: rec.confirmation.clone(),
                application: rec.loan_date.clone(),
                category: "金消日・面談日".to_string(),
            });
        }
    }

    // 担当順マップ作成
    let staff_map: BTreeMap<String, usize> = config
        .staff_order
        .iter()
        .enumerate()
        .map(|(idx, name)| (name.clone(), idx))
        .collect();

    let compare_by_config = |a: &ExportRow, b: &ExportRow| -> std::cmp::Ordering {
        for sk in &config.sort_keys {
            let cmp = match sk.key.as_str() {
                "識別" => a.category.cmp(&b.category),
                "担当" => {
                    let a_rank = staff_map.get(&a.staff).copied().unwrap_or(9999);
                    let b_rank = staff_map.get(&b.staff).copied().unwrap_or(9999);
                    a_rank.cmp(&b_rank)
                }
                "時間" => {
                    let a_sec = parse_time_to_seconds(&a.time);
                    let b_sec = parse_time_to_seconds(&b.time);
                    a_sec.cmp(&b_sec)
                }
                "形態" => a.form_type.cmp(&b.form_type),
                "依頼内容" => a.request_content.cmp(&b.request_content),
                "管轄" => a.jurisdiction.cmp(&b.jurisdiction),
                "お客様氏名" => a.client_name.cmp(&b.client_name),
                _ => std::cmp::Ordering::Equal,
            };

            let final_cmp = if sk.ascending { cmp } else { cmp.reverse() };
            if final_cmp != std::cmp::Ordering::Equal {
                return final_cmp;
            }
        }
        std::cmp::Ordering::Equal
    };

    loan_rows.sort_by(|a, b| a.date.cmp(&b.date).then_with(|| compare_by_config(a, b)));
    cancel_rows.sort_by(|a, b| a.date.cmp(&b.date).then_with(|| compare_by_config(a, b)));

    let mut combined_rows: Vec<ExportRow> = Vec::new();
    combined_rows.extend(loan_rows.clone());
    combined_rows.extend(cancel_rows.clone());
    combined_rows.sort_by(|a, b| a.date.cmp(&b.date).then_with(|| compare_by_config(a, b)));

    let mut date_groups: BTreeMap<String, Vec<ExportRow>> = BTreeMap::new();
    for row in &combined_rows {
        if !row.date.is_empty() {
            date_groups
                .entry(row.date.clone())
                .or_default()
                .push(row.clone());
        }
    }

    for group in date_groups.values_mut() {
        group.sort_by(compare_by_config);
    }

    SortedViewsData {
        loan_rows,
        cancel_rows,
        combined_rows,
        date_groups,
    }
}

/// 書式設定済みの共通フォーマットを生成
fn create_styles() -> (Format, Format) {
    let header_format = Format::new()
        .set_bold()
        .set_border_bottom(FormatBorder::Dotted)
        .set_border_top(FormatBorder::Dotted)
        .set_border_left(FormatBorder::Dotted)
        .set_border_right(FormatBorder::Dotted)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_shrink();

    let data_format = Format::new()
        .set_border_bottom(FormatBorder::Dotted)
        .set_border_top(FormatBorder::Dotted)
        .set_border_left(FormatBorder::Dotted)
        .set_border_right(FormatBorder::Dotted)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_shrink();

    (header_format, data_format)
}

/// 列幅の定義ルール
fn get_col_width(header: &str) -> f64 {
    match header {
        "担当" => 7.0,
        "依頼内容" => 12.0,
        "形態" => 12.0,
        "日付" | "融資実行日" | "金消日・面談日" | "受諾日" => 12.0,
        "時間" | "金消時間" | "立会時間" => 10.0,
        "識別" => 14.0,
        "物件" => 27.0,
        "お客様氏名" => 21.0,
        "場所" | "金消場所・面談場所" | "立会場所" => 25.0,
        "確認者" | "意思確認" | "立会者" => 9.0,
        "申請" | "当日申請" => 11.0,
        "管轄" => 10.0,
        "仲介" => 18.0,
        "先方書士" => 16.0,
        "備考" => 15.0,
        "実行時間" => 10.0,
        _ => 11.5,
    }
}

pub fn export_records_to_excel(
    records: &[Record],
    config: &AppConfig,
    target_dir: &Path,
    file_name: &str,
) -> Result<PathBuf> {
    fs::create_dir_all(target_dir)?;

    let mut workbook = Workbook::new();
    let (header_fmt, data_fmt) = create_styles();
    let row_height = 25.0;

    // 1. 全レコードの「受託確認票」元マスターシート
    let master_headers = vec![
        "ID", "ステータス", "受諾日", "融資実行日", "融資実行日補足", "形態", "お客様氏名", "物件",
        "依頼内容", "金消日・面談日", "金消日補足", "金消場所・面談場所", "金消時間", "担当",
        "意思確認", "申請書類チェック", "ＡＲＵＨＩ代理店", "仲介", "管轄", "先方書士", "備考",
        "立会時間", "立会場所", "立会者", "当日申請", "実行時間", "申請書", "必要書類", "移転書類2",
        "完了書類", "受取日", "東京へ完了", "入金日",
    ];

    let ws_master = workbook.add_worksheet();
    ws_master.set_name("受託確認票")?;

    for (c, h) in master_headers.iter().enumerate() {
        ws_master.set_column_width(c as u16, get_col_width(h))?;
        ws_master.write_string_with_format(0, c as u16, *h, &header_fmt)?;
    }
    ws_master.set_row_height(0, row_height)?;

    for (r, rec) in records.iter().enumerate() {
        let row_idx = (r + 1) as u32;
        ws_master.set_row_height(row_idx, row_height)?;

        let vals = vec![
            rec.id.to_string(),
            rec.status.clone(),
            rec.acceptance_date.clone(),
            rec.loan_date.clone(),
            rec.loan_date_note.clone(),
            rec.form_type.clone(),
            rec.client_name.clone(),
            rec.property_name.clone(),
            rec.request_content.clone(),
            rec.settlement_date.clone(),
            rec.settlement_date_note.clone(),
            rec.settlement_place.clone(),
            rec.settlement_time.clone(),
            rec.staff.clone(),
            rec.confirmation.clone(),
            rec.doc_check.clone(),
            rec.agency.clone(),
            rec.broker.clone(),
            rec.jurisdiction.clone(),
            rec.opponent_scrivener.clone(),
            rec.remarks.clone(),
            rec.attendance_time.clone(),
            rec.attendance_place.clone(),
            rec.attendee.clone(),
            rec.same_day_application.clone(),
            rec.execution_time.clone(),
            rec.progress_application.clone(),
            rec.progress_required_docs.clone(),
            rec.progress_transfer_docs.clone(),
            rec.progress_complete_docs.clone(),
            rec.progress_receive_date.clone(),
            rec.progress_tokyo_complete.clone(),
            rec.progress_payment_date.clone(),
        ];

        for (c, val) in vals.iter().enumerate() {
            ws_master.write_string_with_format(row_idx, c as u16, val, &data_fmt)?;
        }
    }

    // 2. ソート済みデータの構築
    let views = build_sorted_views(records, config);

    let sub_headers = vec![
        "日付", "形態", "お客様氏名", "物件", "依頼内容", "担当", "管轄", "時間", "場所",
        "確認者", "申請", "識別",
    ];

    let write_sub_sheet = |wb: &mut Workbook, sheet_name: &str, rows: &[ExportRow]| -> Result<()> {
        let ws = wb.add_worksheet();
        ws.set_name(sheet_name)?;
        for (c, h) in sub_headers.iter().enumerate() {
            ws.set_column_width(c as u16, get_col_width(h))?;
            ws.write_string_with_format(0, c as u16, *h, &header_fmt)?;
        }
        ws.set_row_height(0, row_height)?;

        for (r, item) in rows.iter().enumerate() {
            let row_idx = (r + 1) as u32;
            ws.set_row_height(row_idx, row_height)?;
            let vals = vec![
                &item.date,
                &item.form_type,
                &item.client_name,
                &item.property_name,
                &item.request_content,
                &item.staff,
                &item.jurisdiction,
                &item.time,
                &item.place,
                &item.confirmer,
                &item.application,
                &item.category,
            ];
            for (c, val) in vals.iter().enumerate() {
                ws.write_string_with_format(row_idx, c as u16, *val, &data_fmt)?;
            }
        }
        Ok(())
    };

    write_sub_sheet(&mut workbook, "sorted_by_融資実行日", &views.loan_rows)?;
    write_sub_sheet(&mut workbook, "sorted_by_金消日・面談日", &views.cancel_rows)?;
    write_sub_sheet(&mut workbook, "統合データ", &views.combined_rows)?;

    // 3. 日付別シートの生成
    for (date_str, group) in views.date_groups {
        let safe_name = date_str.replace('/', "-");
        let sheet_name = if safe_name.len() > 31 {
            &safe_name[..31]
        } else {
            &safe_name
        };

        write_sub_sheet(&mut workbook, sheet_name, &group)?;
    }

    // 4. 一時ファイルへ保存後、アトミックにリネーム・置換（ロック対策）
    let final_path = target_dir.join(file_name);
    let tmp_path = target_dir.join(format!("{}.tmp", file_name));
    let backup_path = target_dir.join(file_name.replace(".xlsx", "_backup.xlsx"));

    workbook.save(&tmp_path)?;

    if final_path.exists() {
        let _ = fs::copy(&final_path, &backup_path);
    }

    let mut replaced = false;
    for _ in 0..5 {
        if fs::rename(&tmp_path, &final_path).is_ok() {
            replaced = true;
            break;
        }
        thread::sleep(Duration::from_millis(500));
    }

    if !replaced {
        if fs::copy(&tmp_path, &final_path).is_err() {
            tracing::warn!("ファイルが他のプロセスで開かれているため即時置換できませんでした。次回更新時に反映されます: {:?}", final_path);
        } else {
            let _ = fs::remove_file(&tmp_path);
        }
    }

    tracing::info!("Excelファイルを正常に出力しました: {:?}", final_path);
    Ok(final_path)
}
