use crate::models::{AppConfig, AuditLog, Record, RecordInput};
use anyhow::{bail, Result};
use chrono::Local;
use rusqlite::{params, Connection};
use std::sync::{Arc, Mutex};

#[derive(Clone)]
pub struct Db {
    conn: Arc<Mutex<Connection>>,
}

impl Db {
    pub fn open(path: &str) -> Result<Self> {
        let conn = Connection::open(path)?;
        let db = Self {
            conn: Arc::new(Mutex::new(conn)),
        };
        db.init_schema()?;
        Ok(db)
    }

    #[allow(dead_code)]
    pub fn open_in_memory() -> Result<Self> {
        let conn = Connection::open_in_memory()?;
        let db = Self {
            conn: Arc::new(Mutex::new(conn)),
        };
        db.init_schema()?;
        Ok(db)
    }

    fn init_schema(&self) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        conn.execute_batch(
            r#"
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                status TEXT NOT NULL DEFAULT '進行中',
                acceptance_date TEXT NOT NULL DEFAULT '',
                loan_date TEXT NOT NULL DEFAULT '',
                loan_date_note TEXT NOT NULL DEFAULT '',
                form_type TEXT NOT NULL DEFAULT '',
                client_name TEXT NOT NULL DEFAULT '',
                property_name TEXT NOT NULL DEFAULT '',
                request_content TEXT NOT NULL DEFAULT '',
                settlement_date TEXT NOT NULL DEFAULT '',
                settlement_date_note TEXT NOT NULL DEFAULT '',
                settlement_place TEXT NOT NULL DEFAULT '',
                settlement_time TEXT NOT NULL DEFAULT '',
                staff TEXT NOT NULL DEFAULT '',
                confirmation TEXT NOT NULL DEFAULT '',
                doc_check TEXT NOT NULL DEFAULT '',
                agency TEXT NOT NULL DEFAULT '',
                broker TEXT NOT NULL DEFAULT '',
                jurisdiction TEXT NOT NULL DEFAULT '',
                opponent_scrivener TEXT NOT NULL DEFAULT '',
                remarks TEXT NOT NULL DEFAULT '',
                attendance_time TEXT NOT NULL DEFAULT '',
                attendance_place TEXT NOT NULL DEFAULT '',
                attendee TEXT NOT NULL DEFAULT '',
                same_day_application TEXT NOT NULL DEFAULT '',
                execution_time TEXT NOT NULL DEFAULT '',
                progress_application TEXT NOT NULL DEFAULT '',
                progress_required_docs TEXT NOT NULL DEFAULT '',
                progress_transfer_docs TEXT NOT NULL DEFAULT '',
                progress_complete_docs TEXT NOT NULL DEFAULT '',
                progress_receive_date TEXT NOT NULL DEFAULT '',
                progress_tokyo_complete TEXT NOT NULL DEFAULT '',
                progress_payment_date TEXT NOT NULL DEFAULT '',
                is_deleted INTEGER NOT NULL DEFAULT 0,
                version INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                client_ip TEXT NOT NULL DEFAULT '',
                actor_name TEXT NOT NULL DEFAULT '',
                action_type TEXT NOT NULL,
                target_id INTEGER,
                summary TEXT NOT NULL DEFAULT '',
                details TEXT NOT NULL DEFAULT ''
            );
            "#,
        )?;

        // 設定の初期値確認
        let mut stmt = conn.prepare("SELECT value FROM settings WHERE key = 'app_config'")?;
        let exists = stmt.exists([])?;
        drop(stmt);

        if !exists {
            let default_cfg = AppConfig::default();
            let json = serde_json::to_string(&default_cfg)?;
            conn.execute(
                "INSERT INTO settings (key, value) VALUES ('app_config', ?1)",
                params![json],
            )?;
        }

        Ok(())
    }

    pub fn get_config(&self) -> Result<AppConfig> {
        let conn = self.conn.lock().unwrap();
        let mut stmt = conn.prepare("SELECT value FROM settings WHERE key = 'app_config'")?;
        let json: String = stmt.query_row([], |row| row.get(0))?;
        let cfg: AppConfig = serde_json::from_str(&json)?;
        Ok(cfg)
    }

    pub fn update_config(&self, cfg: &AppConfig) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        let json = serde_json::to_string(cfg)?;
        conn.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES ('app_config', ?1)",
            params![json],
        )?;
        Ok(())
    }

    pub fn list_records(
        &self,
        query: Option<&str>,
        staff_filter: Option<&str>,
        status_filter: Option<&str>,
        include_deleted: bool,
    ) -> Result<Vec<Record>> {
        let conn = self.conn.lock().unwrap();
        let mut sql = String::from("SELECT id, status, acceptance_date, loan_date, loan_date_note, form_type, client_name, property_name, request_content, settlement_date, settlement_date_note, settlement_place, settlement_time, staff, confirmation, doc_check, agency, broker, jurisdiction, opponent_scrivener, remarks, attendance_time, attendance_place, attendee, same_day_application, execution_time, progress_application, progress_required_docs, progress_transfer_docs, progress_complete_docs, progress_receive_date, progress_tokyo_complete, progress_payment_date, is_deleted, version, created_at, updated_at FROM records WHERE 1=1");

        let mut param_values: Vec<String> = Vec::new();

        if !include_deleted {
            sql.push_str(" AND is_deleted = 0");
        }

        if let Some(s) = staff_filter {
            if !s.is_empty() {
                sql.push_str(" AND staff = ?");
                param_values.push(s.to_string());
            }
        }

        if let Some(st) = status_filter {
            if !st.is_empty() {
                sql.push_str(" AND status = ?");
                param_values.push(st.to_string());
            }
        }

        if let Some(q) = query {
            let q = q.trim();
            if !q.is_empty() {
                sql.push_str(" AND (client_name LIKE ? OR property_name LIKE ? OR broker LIKE ? OR remarks LIKE ?)");
                let pattern = format!("%{}%", q);
                param_values.push(pattern.clone());
                param_values.push(pattern.clone());
                param_values.push(pattern.clone());
                param_values.push(pattern);
            }
        }

        sql.push_str(" ORDER BY id DESC");

        let mut stmt = conn.prepare(&sql)?;
        let rows = stmt.query_map(rusqlite::params_from_iter(param_values.iter()), |row| {
            Ok(Record {
                id: row.get(0)?,
                status: row.get(1)?,
                acceptance_date: row.get(2)?,
                loan_date: row.get(3)?,
                loan_date_note: row.get(4)?,
                form_type: row.get(5)?,
                client_name: row.get(6)?,
                property_name: row.get(7)?,
                request_content: row.get(8)?,
                settlement_date: row.get(9)?,
                settlement_date_note: row.get(10)?,
                settlement_place: row.get(11)?,
                settlement_time: row.get(12)?,
                staff: row.get(13)?,
                confirmation: row.get(14)?,
                doc_check: row.get(15)?,
                agency: row.get(16)?,
                broker: row.get(17)?,
                jurisdiction: row.get(18)?,
                opponent_scrivener: row.get(19)?,
                remarks: row.get(20)?,
                attendance_time: row.get(21)?,
                attendance_place: row.get(22)?,
                attendee: row.get(23)?,
                same_day_application: row.get(24)?,
                execution_time: row.get(25)?,
                progress_application: row.get(26)?,
                progress_required_docs: row.get(27)?,
                progress_transfer_docs: row.get(28)?,
                progress_complete_docs: row.get(29)?,
                progress_receive_date: row.get(30)?,
                progress_tokyo_complete: row.get(31)?,
                progress_payment_date: row.get(32)?,
                is_deleted: row.get(33)?,
                version: row.get(34)?,
                created_at: row.get(35)?,
                updated_at: row.get(36)?,
            })
        })?;

        let mut records = Vec::new();
        for r in rows {
            records.push(r?);
        }
        Ok(records)
    }

    pub fn get_record(&self, id: i64) -> Result<Option<Record>> {
        let conn = self.conn.lock().unwrap();
        let mut stmt = conn.prepare("SELECT id, status, acceptance_date, loan_date, loan_date_note, form_type, client_name, property_name, request_content, settlement_date, settlement_date_note, settlement_place, settlement_time, staff, confirmation, doc_check, agency, broker, jurisdiction, opponent_scrivener, remarks, attendance_time, attendance_place, attendee, same_day_application, execution_time, progress_application, progress_required_docs, progress_transfer_docs, progress_complete_docs, progress_receive_date, progress_tokyo_complete, progress_payment_date, is_deleted, version, created_at, updated_at FROM records WHERE id = ?1")?;

        let mut rows = stmt.query(params![id])?;
        if let Some(row) = rows.next()? {
            Ok(Some(Record {
                id: row.get(0)?,
                status: row.get(1)?,
                acceptance_date: row.get(2)?,
                loan_date: row.get(3)?,
                loan_date_note: row.get(4)?,
                form_type: row.get(5)?,
                client_name: row.get(6)?,
                property_name: row.get(7)?,
                request_content: row.get(8)?,
                settlement_date: row.get(9)?,
                settlement_date_note: row.get(10)?,
                settlement_place: row.get(11)?,
                settlement_time: row.get(12)?,
                staff: row.get(13)?,
                confirmation: row.get(14)?,
                doc_check: row.get(15)?,
                agency: row.get(16)?,
                broker: row.get(17)?,
                jurisdiction: row.get(18)?,
                opponent_scrivener: row.get(19)?,
                remarks: row.get(20)?,
                attendance_time: row.get(21)?,
                attendance_place: row.get(22)?,
                attendee: row.get(23)?,
                same_day_application: row.get(24)?,
                execution_time: row.get(25)?,
                progress_application: row.get(26)?,
                progress_required_docs: row.get(27)?,
                progress_transfer_docs: row.get(28)?,
                progress_complete_docs: row.get(29)?,
                progress_receive_date: row.get(30)?,
                progress_tokyo_complete: row.get(31)?,
                progress_payment_date: row.get(32)?,
                is_deleted: row.get(33)?,
                version: row.get(34)?,
                created_at: row.get(35)?,
                updated_at: row.get(36)?,
            }))
        } else {
            Ok(None)
        }
    }

    pub fn insert_record(&self, input: &RecordInput) -> Result<i64> {
        let conn = self.conn.lock().unwrap();
        let now = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();

        conn.execute(
            r#"
            INSERT INTO records (
                status, acceptance_date, loan_date, loan_date_note, form_type, client_name, property_name,
                request_content, settlement_date, settlement_date_note, settlement_place, settlement_time,
                staff, confirmation, doc_check, agency, broker, jurisdiction, opponent_scrivener,
                remarks, attendance_time, attendance_place, attendee, same_day_application, execution_time,
                progress_application, progress_required_docs, progress_transfer_docs, progress_complete_docs,
                progress_receive_date, progress_tokyo_complete, progress_payment_date, is_deleted, version,
                created_at, updated_at
            ) VALUES (
                ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19,
                ?20, ?21, ?22, ?23, ?24, ?25, ?26, ?27, ?28, ?29, ?30, ?31, ?32, 0, 1, ?33, ?34
            )
            "#,
            params![
                input.status.as_deref().unwrap_or("進行中"),
                input.acceptance_date.as_deref().unwrap_or(""),
                input.loan_date.as_deref().unwrap_or(""),
                input.loan_date_note.as_deref().unwrap_or(""),
                input.form_type.as_deref().unwrap_or(""),
                input.client_name.as_deref().unwrap_or(""),
                input.property_name.as_deref().unwrap_or(""),
                input.request_content.as_deref().unwrap_or(""),
                input.settlement_date.as_deref().unwrap_or(""),
                input.settlement_date_note.as_deref().unwrap_or(""),
                input.settlement_place.as_deref().unwrap_or(""),
                input.settlement_time.as_deref().unwrap_or(""),
                input.staff.as_deref().unwrap_or(""),
                input.confirmation.as_deref().unwrap_or(""),
                input.doc_check.as_deref().unwrap_or(""),
                input.agency.as_deref().unwrap_or(""),
                input.broker.as_deref().unwrap_or(""),
                input.jurisdiction.as_deref().unwrap_or(""),
                input.opponent_scrivener.as_deref().unwrap_or(""),
                input.remarks.as_deref().unwrap_or(""),
                input.attendance_time.as_deref().unwrap_or(""),
                input.attendance_place.as_deref().unwrap_or(""),
                input.attendee.as_deref().unwrap_or(""),
                input.same_day_application.as_deref().unwrap_or(""),
                input.execution_time.as_deref().unwrap_or(""),
                input.progress_application.as_deref().unwrap_or(""),
                input.progress_required_docs.as_deref().unwrap_or(""),
                input.progress_transfer_docs.as_deref().unwrap_or(""),
                input.progress_complete_docs.as_deref().unwrap_or(""),
                input.progress_receive_date.as_deref().unwrap_or(""),
                input.progress_tokyo_complete.as_deref().unwrap_or(""),
                input.progress_payment_date.as_deref().unwrap_or(""),
                now,
                now
            ],
        )?;

        Ok(conn.last_insert_rowid())
    }

    pub fn update_record(&self, id: i64, input: &RecordInput) -> Result<()> {
        let conn = self.conn.lock().unwrap();

        // 楽観的ロックのチェック
        let mut stmt = conn.prepare("SELECT version FROM records WHERE id = ?1 AND is_deleted = 0")?;
        let current_ver: Result<i64, _> = stmt.query_row(params![id], |row| row.get(0));
        match current_ver {
            Ok(v) => {
                if let Some(expected_ver) = input.version {
                    if v != expected_ver {
                        bail!("CONFLICT: 別のユーザーによってこの案件は既に更新されています。最新情報を読み込み直してください。");
                    }
                }
            }
            Err(_) => {
                bail!("NOT_FOUND: 指定された案件が見つかりません。");
            }
        }

        let now = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();

        let updated = conn.execute(
            r#"
            UPDATE records SET
                status = coalesce(?1, status),
                acceptance_date = coalesce(?2, acceptance_date),
                loan_date = coalesce(?3, loan_date),
                loan_date_note = coalesce(?4, loan_date_note),
                form_type = coalesce(?5, form_type),
                client_name = coalesce(?6, client_name),
                property_name = coalesce(?7, property_name),
                request_content = coalesce(?8, request_content),
                settlement_date = coalesce(?9, settlement_date),
                settlement_date_note = coalesce(?10, settlement_date_note),
                settlement_place = coalesce(?11, settlement_place),
                settlement_time = coalesce(?12, settlement_time),
                staff = coalesce(?13, staff),
                confirmation = coalesce(?14, confirmation),
                doc_check = coalesce(?15, doc_check),
                agency = coalesce(?16, agency),
                broker = coalesce(?17, broker),
                jurisdiction = coalesce(?18, jurisdiction),
                opponent_scrivener = coalesce(?19, opponent_scrivener),
                remarks = coalesce(?20, remarks),
                attendance_time = coalesce(?21, attendance_time),
                attendance_place = coalesce(?22, attendance_place),
                attendee = coalesce(?23, attendee),
                same_day_application = coalesce(?24, same_day_application),
                execution_time = coalesce(?25, execution_time),
                progress_application = coalesce(?26, progress_application),
                progress_required_docs = coalesce(?27, progress_required_docs),
                progress_transfer_docs = coalesce(?28, progress_transfer_docs),
                progress_complete_docs = coalesce(?29, progress_complete_docs),
                progress_receive_date = coalesce(?30, progress_receive_date),
                progress_tokyo_complete = coalesce(?31, progress_tokyo_complete),
                progress_payment_date = coalesce(?32, progress_payment_date),
                version = version + 1,
                updated_at = ?33
            WHERE id = ?34 AND is_deleted = 0
            "#,
            params![
                input.status,
                input.acceptance_date,
                input.loan_date,
                input.loan_date_note,
                input.form_type,
                input.client_name,
                input.property_name,
                input.request_content,
                input.settlement_date,
                input.settlement_date_note,
                input.settlement_place,
                input.settlement_time,
                input.staff,
                input.confirmation,
                input.doc_check,
                input.agency,
                input.broker,
                input.jurisdiction,
                input.opponent_scrivener,
                input.remarks,
                input.attendance_time,
                input.attendance_place,
                input.attendee,
                input.same_day_application,
                input.execution_time,
                input.progress_application,
                input.progress_required_docs,
                input.progress_transfer_docs,
                input.progress_complete_docs,
                input.progress_receive_date,
                input.progress_tokyo_complete,
                input.progress_payment_date,
                now,
                id
            ],
        )?;

        if updated == 0 {
            bail!("NOT_FOUND: 更新対象の案件が見つかりません。");
        }

        Ok(())
    }

    pub fn delete_record(&self, id: i64) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        let now = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
        conn.execute(
            "UPDATE records SET is_deleted = 1, updated_at = ?1 WHERE id = ?2",
            params![now, id],
        )?;
        Ok(())
    }

    /// 論理削除されたレコード一覧を取得（ゴミ箱）
    pub fn list_deleted_records(&self) -> Result<Vec<Record>> {
        let conn = self.conn.lock().unwrap();
        let sql = "SELECT id, status, acceptance_date, loan_date, loan_date_note, form_type, client_name, property_name, request_content, settlement_date, settlement_date_note, settlement_place, settlement_time, staff, confirmation, doc_check, agency, broker, jurisdiction, opponent_scrivener, remarks, attendance_time, attendance_place, attendee, same_day_application, execution_time, progress_application, progress_required_docs, progress_transfer_docs, progress_complete_docs, progress_receive_date, progress_tokyo_complete, progress_payment_date, is_deleted, version, created_at, updated_at FROM records WHERE is_deleted = 1 ORDER BY updated_at DESC, id DESC";
        let mut stmt = conn.prepare(sql)?;
        let rows = stmt.query_map([], |row| {
            Ok(Record {
                id: row.get(0)?,
                status: row.get(1)?,
                acceptance_date: row.get(2)?,
                loan_date: row.get(3)?,
                loan_date_note: row.get(4)?,
                form_type: row.get(5)?,
                client_name: row.get(6)?,
                property_name: row.get(7)?,
                request_content: row.get(8)?,
                settlement_date: row.get(9)?,
                settlement_date_note: row.get(10)?,
                settlement_place: row.get(11)?,
                settlement_time: row.get(12)?,
                staff: row.get(13)?,
                confirmation: row.get(14)?,
                doc_check: row.get(15)?,
                agency: row.get(16)?,
                broker: row.get(17)?,
                jurisdiction: row.get(18)?,
                opponent_scrivener: row.get(19)?,
                remarks: row.get(20)?,
                attendance_time: row.get(21)?,
                attendance_place: row.get(22)?,
                attendee: row.get(23)?,
                same_day_application: row.get(24)?,
                execution_time: row.get(25)?,
                progress_application: row.get(26)?,
                progress_required_docs: row.get(27)?,
                progress_transfer_docs: row.get(28)?,
                progress_complete_docs: row.get(29)?,
                progress_receive_date: row.get(30)?,
                progress_tokyo_complete: row.get(31)?,
                progress_payment_date: row.get(32)?,
                is_deleted: row.get(33)?,
                version: row.get(34)?,
                created_at: row.get(35)?,
                updated_at: row.get(36)?,
            })
        })?;

        let mut records = Vec::new();
        for r in rows {
            records.push(r?);
        }
        Ok(records)
    }

    /// 論理削除されたレコードを復元
    pub fn restore_record(&self, id: i64) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        let now = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
        conn.execute(
            "UPDATE records SET is_deleted = 0, updated_at = ?1 WHERE id = ?2",
            params![now, id],
        )?;
        Ok(())
    }

    /// レコードを物理削除（完全削除）
    pub fn purge_record(&self, id: i64) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        conn.execute("DELETE FROM records WHERE id = ?1", params![id])?;
        Ok(())
    }

    /// 全レコードをクリア（インポート時の置換モード用）
    pub fn clear_all_records(&self) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        conn.execute("DELETE FROM records", [])?;
        Ok(())
    }

    /// バックアップDBファイルから現在のDBへデータを復元
    pub fn restore_from_backup_file(&self, backup_db_path: &std::path::Path) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        let path_str = backup_db_path.to_string_lossy().replace('\'', "''");
        conn.execute_batch(&format!(
            "ATTACH DATABASE '{}' AS backup_db;
             BEGIN;
             DELETE FROM main.records;
             INSERT INTO main.records SELECT * FROM backup_db.records;
             COMMIT;
             DETACH DATABASE backup_db;",
            path_str
        ))?;
        Ok(())
    }

    /// データベースの全有効レコード数を取得
    pub fn count_records(&self) -> Result<i64> {
        let conn = self.conn.lock().unwrap();
        let count: i64 = conn.query_row(
            "SELECT count(*) FROM records WHERE is_deleted = 0",
            [],
            |r| r.get(0),
        )?;
        Ok(count)
    }

    /// 監査ログ（アクセス・操作記録）を保存
    pub fn insert_audit_log(
        &self,
        client_ip: &str,
        actor_name: &str,
        action_type: &str,
        target_id: Option<i64>,
        summary: &str,
        details: &str,
    ) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        let now = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
        conn.execute(
            r#"
            INSERT INTO audit_logs (created_at, client_ip, actor_name, action_type, target_id, summary, details)
            VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)
            "#,
            params![now, client_ip, actor_name, action_type, target_id, summary, details],
        )?;
        Ok(())
    }

    /// 監査ログ一覧を取得（最新順、最大limit件）
    pub fn list_audit_logs(
        &self,
        action_filter: Option<&str>,
        query: Option<&str>,
        limit: usize,
    ) -> Result<Vec<AuditLog>> {
        let conn = self.conn.lock().unwrap();
        let mut sql = String::from(
            "SELECT id, created_at, client_ip, actor_name, action_type, target_id, summary, details FROM audit_logs WHERE 1=1",
        );
        let mut param_values: Vec<String> = Vec::new();

        if let Some(af) = action_filter {
            if !af.is_empty() {
                sql.push_str(" AND action_type = ?");
                param_values.push(af.to_string());
            }
        }

        if let Some(q) = query {
            let q = q.trim();
            if !q.is_empty() {
                sql.push_str(" AND (client_ip LIKE ? OR actor_name LIKE ? OR summary LIKE ? OR details LIKE ?)");
                let pattern = format!("%{}%", q);
                param_values.push(pattern.clone());
                param_values.push(pattern.clone());
                param_values.push(pattern.clone());
                param_values.push(pattern);
            }
        }

        sql.push_str(&format!(" ORDER BY id DESC LIMIT {}", limit.min(2000)));

        let mut stmt = conn.prepare(&sql)?;
        let rows = stmt.query_map(rusqlite::params_from_iter(param_values.iter()), |row| {
            Ok(AuditLog {
                id: row.get(0)?,
                created_at: row.get(1)?,
                client_ip: row.get(2)?,
                actor_name: row.get(3)?,
                action_type: row.get(4)?,
                target_id: row.get(5)?,
                summary: row.get(6)?,
                details: row.get(7)?,
            })
        })?;

        let mut logs = Vec::new();
        for r in rows {
            logs.push(r?);
        }
        Ok(logs)
    }

    /// 指定IDの監査ログを削除
    pub fn delete_audit_log(&self, id: i64) -> Result<()> {
        let conn = self.conn.lock().unwrap();
        conn.execute("DELETE FROM audit_logs WHERE id = ?1", params![id])?;
        Ok(())
    }

    /// テスト用または全監査ログのクリア
    pub fn clear_audit_logs(&self, only_test: bool) -> Result<usize> {
        let conn = self.conn.lock().unwrap();
        let count = if only_test {
            conn.execute(
                "DELETE FROM audit_logs WHERE actor_name LIKE '%テスト%' OR summary LIKE '%テスト%'",
                [],
            )?
        } else {
            conn.execute("DELETE FROM audit_logs", [])?
        };
        Ok(count)
    }
}
