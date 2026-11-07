use anyhow::Result;
use chrono::{Duration, Local, NaiveDate};
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BackupEntry {
    pub date_str: String,
    pub excel_file: Option<String>,
    pub db_file: Option<String>,
    pub created_at: String,
}

pub struct BackupManager {
    backup_dir: PathBuf,
    retention_days: u32,
}

impl BackupManager {
    pub fn new(backup_dir: impl AsRef<Path>, retention_days: u32) -> Self {
        Self {
            backup_dir: backup_dir.as_ref().to_path_buf(),
            retention_days,
        }
    }

    /// 日次バックアップを実行
    pub fn create_daily_backup(&self, current_excel_path: &Path, current_db_path: &Path) -> Result<()> {
        fs::create_dir_all(&self.backup_dir)?;
        let today = Local::now().format("%Y.%m.%d").to_string();

        // 1. Excelのバックアップ: YYYY.MM.DD 受託確認票.xlsx
        if current_excel_path.exists() {
            let backup_excel_name = format!("{} 受託確認票.xlsx", today);
            let backup_excel_path = self.backup_dir.join(&backup_excel_name);
            fs::copy(current_excel_path, &backup_excel_path)?;
            tracing::info!("Excel日次バックアップを作成しました: {:?}", backup_excel_path);
        }

        // 2. データベースのバックアップ: YYYY.MM.DD_data.db
        if current_db_path.exists() {
            let backup_db_name = format!("{}_data.db", today);
            let backup_db_path = self.backup_dir.join(&backup_db_name);
            fs::copy(current_db_path, &backup_db_path)?;
            tracing::info!("DB日次バックアップを作成しました: {:?}", backup_db_path);
        }

        // 3. 古いバックアップのローテーションクリーンアップ
        self.cleanup_old_backups()?;

        Ok(())
    }

    /// 期限切れバックアップの自動削除
    pub fn cleanup_old_backups(&self) -> Result<()> {
        if !self.backup_dir.exists() {
            return Ok(());
        }

        let threshold = Local::now().date_naive() - Duration::days(self.retention_days as i64);

        for entry in fs::read_dir(&self.backup_dir)? {
            let entry = entry?;
            let path = entry.path();
            if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
                // ファイル名先頭の YYYY.MM.DD をパース
                if file_name.len() >= 10 {
                    let date_part = &file_name[..10];
                    if let Ok(file_date) = NaiveDate::parse_from_str(date_part, "%Y.%m.%d") {
                        if file_date < threshold {
                            let _ = fs::remove_file(&path);
                            tracing::info!("保持期間を超えた古いバックアップを削除しました: {:?}", path);
                        }
                    }
                }
            }
        }
        Ok(())
    }

    /// 現在のバックアップ一覧を取得
    pub fn list_backups(&self) -> Result<Vec<BackupEntry>> {
        if !self.backup_dir.exists() {
            return Ok(Vec::new());
        }

        let mut map: std::collections::BTreeMap<String, (Option<String>, Option<String>, String)> =
            std::collections::BTreeMap::new();

        for entry in fs::read_dir(&self.backup_dir)? {
            let entry = entry?;
            let path = entry.path();
            if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
                if file_name.len() >= 10 {
                    let date_part = &file_name[..10];
                    if NaiveDate::parse_from_str(date_part, "%Y.%m.%d").is_ok() {
                        let metadata = fs::metadata(&path)?;
                        let mod_time = metadata
                            .modified()
                            .ok()
                            .and_then(|t| {
                                let dt: chrono::DateTime<Local> = t.into();
                                Some(dt.format("%Y-%m-%d %H:%M:%S").to_string())
                            })
                            .unwrap_or_default();

                        let item = map
                            .entry(date_part.to_string())
                            .or_insert((None, None, mod_time));
                        if file_name.ends_with(".xlsx") {
                            item.0 = Some(file_name.to_string());
                        } else if file_name.ends_with(".db") {
                            item.1 = Some(file_name.to_string());
                        }
                    }
                }
            }
        }

        let mut result = Vec::new();
        for (date_str, (excel, db, time)) in map.into_iter().rev() {
            result.push(BackupEntry {
                date_str,
                excel_file: excel,
                db_file: db,
                created_at: time,
            });
        }

        Ok(result)
    }
}
