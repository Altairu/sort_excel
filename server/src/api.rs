use crate::backup::BackupManager;
use crate::db::Db;
use crate::excel_exporter::export_records_to_excel;
use crate::excel_importer::import_excel_file;
use crate::models::{AppConfig, RecordInput};
use axum::{
    extract::{ConnectInfo, Path, Query, State},
    http::StatusCode,
    response::IntoResponse,
    Json,
};
use serde::{Deserialize, Serialize};
use std::net::SocketAddr;
use std::path::{Path as StdPath, PathBuf};
use std::sync::{Arc, Mutex};

pub struct AppState {
    pub db: Db,
    pub db_path: PathBuf,
    pub backup_dir: PathBuf,
    pub export_lock: Arc<Mutex<()>>,
}

#[derive(Debug, Deserialize)]
pub struct ListParams {
    pub q: Option<String>,
    pub staff: Option<String>,
    pub status: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct ApiResponse<T> {
    pub success: bool,
    pub data: Option<T>,
    pub error: Option<String>,
}

impl<T> ApiResponse<T> {
    pub fn ok(data: T) -> Self {
        Self {
            success: true,
            data: Some(data),
            error: None,
        }
    }
    pub fn err(msg: impl Into<String>) -> Self {
        Self {
            success: false,
            data: None,
            error: Some(msg.into()),
        }
    }
}

/// 接続元がサーバーPC自身（ループバックアドレス）かどうかを判定
pub fn is_local_request(addr: &SocketAddr) -> bool {
    addr.ip().is_loopback()
}

/// 接続元情報を取得（使用者画面で管理者リンクを表示するかどうかの判定用）
#[derive(Debug, Serialize)]
pub struct SessionInfo {
    pub is_admin_pc: bool,
}

pub async fn get_session_info(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
) -> impl IntoResponse {
    (
        StatusCode::OK,
        Json(ApiResponse::ok(SessionInfo {
            is_admin_pc: is_local_request(&addr),
        })),
    )
}

/// 自動エクスポートヘルパー（ミューテックスで同時書き込み競合を防止）
fn trigger_auto_export(state: &AppState) {
    let db = state.db.clone();
    let export_lock = state.export_lock.clone();

    tokio::task::spawn_blocking(move || {
        let _guard = match export_lock.lock() {
            Ok(g) => g,
            Err(poisoned) => poisoned.into_inner(),
        };
        if let Ok(config) = db.get_config() {
            if config.auto_export_on_change {
                if let Ok(records) = db.list_records(None, None, None, false) {
                    let export_dir = PathBuf::from(&config.export_dir);
                    if let Err(e) =
                        export_records_to_excel(&records, &config, &export_dir, &config.export_file_name)
                    {
                        tracing::error!("自動Excel出力エラー: {:?}", e);
                    }
                }
            }
        }
    });
}

/// 案件一覧取得
pub async fn list_records(
    State(state): State<Arc<AppState>>,
    Query(params): Query<ListParams>,
) -> impl IntoResponse {
    match state.db.list_records(
        params.q.as_deref(),
        params.staff.as_deref(),
        params.status.as_deref(),
        false,
    ) {
        Ok(records) => (StatusCode::OK, Json(ApiResponse::ok(records))),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 案件単一取得
pub async fn get_record(
    State(state): State<Arc<AppState>>,
    Path(id): Path<i64>,
) -> impl IntoResponse {
    match state.db.get_record(id) {
        Ok(Some(rec)) => (StatusCode::OK, Json(ApiResponse::ok(rec))),
        Ok(None) => (
            StatusCode::NOT_FOUND,
            Json(ApiResponse::err("案件が見つかりません")),
        ),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 新規案件登録
pub async fn create_record(
    State(state): State<Arc<AppState>>,
    Json(input): Json<RecordInput>,
) -> impl IntoResponse {
    match state.db.insert_record(&input) {
        Ok(id) => {
            trigger_auto_export(&state);
            (StatusCode::CREATED, Json(ApiResponse::ok(id)))
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 案件更新
pub async fn update_record(
    State(state): State<Arc<AppState>>,
    Path(id): Path<i64>,
    Json(input): Json<RecordInput>,
) -> impl IntoResponse {
    match state.db.update_record(id, &input) {
        Ok(()) => {
            trigger_auto_export(&state);
            (StatusCode::OK, Json(ApiResponse::ok("更新完了")))
        }
        Err(e) => {
            let msg = e.to_string();
            let status = if msg.starts_with("CONFLICT") {
                StatusCode::CONFLICT
            } else if msg.starts_with("NOT_FOUND") {
                StatusCode::NOT_FOUND
            } else {
                StatusCode::INTERNAL_SERVER_ERROR
            };
            (status, Json(ApiResponse::err(msg)))
        }
    }
}

/// 案件削除 (論理削除)
pub async fn delete_record(
    State(state): State<Arc<AppState>>,
    Path(id): Path<i64>,
) -> impl IntoResponse {
    match state.db.delete_record(id) {
        Ok(()) => {
            trigger_auto_export(&state);
            (StatusCode::OK, Json(ApiResponse::ok("削除完了")))
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 設定取得（選択肢表示等のため全クライアント参照可）
pub async fn get_config(State(state): State<Arc<AppState>>) -> impl IntoResponse {
    match state.db.get_config() {
        Ok(cfg) => (StatusCode::OK, Json(ApiResponse::ok(cfg))),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 設定更新（管理者PCのみ許可 & パス検証付き）
pub async fn update_config(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Json(mut cfg): Json<AppConfig>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "設定の変更はサーバーPC（管理者画面）からのみ実行できます",
            )),
        );
    }

    // ファイル名の安全性検証（パストラバーサルや拡張子偽装を防止）
    let file_name = cfg.export_file_name.trim().to_string();
    if file_name.is_empty()
        || file_name.contains("..")
        || file_name.contains('/')
        || file_name.contains('\\')
        || file_name.contains(':')
        || !file_name.ends_with(".xlsx")
    {
        return (
            StatusCode::BAD_REQUEST,
            Json(ApiResponse::err(
                "出力ファイル名はパス記号を含まない .xlsx 形式の名前を指定してください",
            )),
        );
    }
    cfg.export_file_name = file_name;

    // 出力先フォルダの検証
    let export_dir = cfg.export_dir.trim().to_string();
    if export_dir.is_empty() || export_dir.contains("..") {
        return (
            StatusCode::BAD_REQUEST,
            Json(ApiResponse::err("不正な出力先フォルダパスです")),
        );
    }
    cfg.export_dir = export_dir;

    if cfg.backup_retention_days == 0 || cfg.backup_retention_days > 365 {
        cfg.backup_retention_days = 7;
    }

    match state.db.update_config(&cfg) {
        Ok(()) => {
            trigger_auto_export(&state);
            (StatusCode::OK, Json(ApiResponse::ok("設定保存完了")))
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 手動Excelエクスポート
pub async fn trigger_export(State(state): State<Arc<AppState>>) -> impl IntoResponse {
    let _guard = match state.export_lock.lock() {
        Ok(g) => g,
        Err(poisoned) => poisoned.into_inner(),
    };

    match state.db.get_config() {
        Ok(cfg) => match state.db.list_records(None, None, None, false) {
            Ok(records) => {
                let export_dir = PathBuf::from(&cfg.export_dir);
                match export_records_to_excel(&records, &cfg, &export_dir, &cfg.export_file_name) {
                    Ok(path) => (
                        StatusCode::OK,
                        Json(ApiResponse::ok(format!(
                            "Excelファイルを出力しました: {:?}",
                            path
                        ))),
                    ),
                    Err(e) => (
                        StatusCode::INTERNAL_SERVER_ERROR,
                        Json(ApiResponse::err(format!("出力失敗: {}", e))),
                    ),
                }
            }
            Err(e) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(ApiResponse::err(e.to_string())),
            ),
        },
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

#[derive(Debug, Deserialize)]
pub struct ImportRequest {
    pub file_path: Option<String>,
    pub mode: Option<String>, // "append" (追加) または "replace" (既存クリアして置換)
}

/// 既存Excelファイルのインポート（管理者PCのみ許可）
pub async fn import_records(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Json(payload): Json<ImportRequest>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "Excel一括インポートはサーバーPC（管理者画面）からのみ実行できます",
            )),
        );
    }

    let path_str = payload
        .file_path
        .unwrap_or_else(|| r"Sample data\2026.10.20 受託確認票.xlsx".to_string());

    if !path_str.to_lowercase().ends_with(".xlsx") {
        return (
            StatusCode::BAD_REQUEST,
            Json(ApiResponse::err(
                "取り込み対象は .xlsx ファイルのみ指定可能です",
            )),
        );
    }

    let path = StdPath::new(&path_str);
    if !path.exists() {
        return (
            StatusCode::BAD_REQUEST,
            Json(ApiResponse::err(format!(
                "指定されたファイルが見つかりません: {:?}",
                path
            ))),
        );
    }

    let is_replace = payload.mode.as_deref() == Some("replace");
    if is_replace {
        if let Err(e) = state.db.clear_all_records() {
            return (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(ApiResponse::err(format!("既存データクリア失敗: {}", e))),
            );
        }
    }

    match import_excel_file(path, &state.db) {
        Ok(count) => {
            trigger_auto_export(&state);
            let mode_label = if is_replace { "総入れ替えで" } else { "追加で" };
            (
                StatusCode::OK,
                Json(ApiResponse::ok(format!(
                    "{}{} 件の案件を取り込みました",
                    mode_label, count
                ))),
            )
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(format!("インポートエラー: {}", e))),
        ),
    }
}

/// バックアップ一覧取得（管理者PCのみ許可）
pub async fn list_backups(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "バックアップ管理はサーバーPC（管理者画面）からのみアクセスできます",
            )),
        );
    }

    let cfg = state.db.get_config().unwrap_or_default();
    let mgr = BackupManager::new(&state.backup_dir, cfg.backup_retention_days);
    match mgr.list_backups() {
        Ok(list) => (StatusCode::OK, Json(ApiResponse::ok(list))),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 今すぐ日次バックアップを手動作成（管理者PCのみ許可）
pub async fn create_backup_now(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "バックアップ作成はサーバーPC（管理者画面）からのみ実行できます",
            )),
        );
    }

    let cfg = state.db.get_config().unwrap_or_default();
    let mgr = BackupManager::new(&state.backup_dir, cfg.backup_retention_days);
    let excel_path = StdPath::new(&cfg.export_dir).join(&cfg.export_file_name);
    match mgr.create_daily_backup(&excel_path, &state.db_path) {
        Ok(()) => (
            StatusCode::OK,
            Json(ApiResponse::ok("本日分のバックアップを作成しました")),
        ),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(format!("バックアップ作成失敗: {}", e))),
        ),
    }
}

#[derive(Debug, Deserialize)]
pub struct RestoreBackupRequest {
    pub date_str: String,
}

/// 指定日付のバックアップからデータを復元（管理者PCのみ許可）
pub async fn restore_backup(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Json(payload): Json<RestoreBackupRequest>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "バックアップ復元はサーバーPC（管理者画面）からのみ実行できます",
            )),
        );
    }

    let date_str = payload.date_str.trim();
    if date_str.len() != 10
        || date_str.contains("..")
        || date_str.contains('/')
        || date_str.contains('\\')
    {
        return (
            StatusCode::BAD_REQUEST,
            Json(ApiResponse::err("不正な日付指定です")),
        );
    }

    let backup_db_path = state.backup_dir.join(format!("{}_data.db", date_str));
    if !backup_db_path.exists() {
        return (
            StatusCode::NOT_FOUND,
            Json(ApiResponse::err(format!(
                "指定日のバックアップDBファイルが見つかりません: {:?}",
                backup_db_path
            ))),
        );
    }

    match state.db.restore_from_backup_file(&backup_db_path) {
        Ok(()) => {
            trigger_auto_export(&state);
            (
                StatusCode::OK,
                Json(ApiResponse::ok(format!(
                    "{} のバックアップからデータを復元し、Excelを更新しました",
                    date_str
                ))),
            )
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(format!("復元エラー: {}", e))),
        ),
    }
}

/// 削除済み案件（ゴミ箱）一覧取得（管理者PCのみ許可）
pub async fn list_trash(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "ゴミ箱の管理はサーバーPC（管理者画面）からのみアクセスできます",
            )),
        );
    }

    match state.db.list_deleted_records() {
        Ok(records) => (StatusCode::OK, Json(ApiResponse::ok(records))),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 削除済み案件を復元（管理者PCのみ許可）
pub async fn restore_trash_record(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Path(id): Path<i64>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "案件の復元はサーバーPC（管理者画面）からのみ実行できます",
            )),
        );
    }

    match state.db.restore_record(id) {
        Ok(()) => {
            trigger_auto_export(&state);
            (StatusCode::OK, Json(ApiResponse::ok("案件を復元しました")))
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 削除済み案件を完全削除（管理者PCのみ許可）
pub async fn purge_trash_record(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Path(id): Path<i64>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "完全削除はサーバーPC（管理者画面）からのみ実行できます",
            )),
        );
    }

    match state.db.purge_record(id) {
        Ok(()) => (
            StatusCode::OK,
            Json(ApiResponse::ok("案件を完全に削除しました")),
        ),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 統計情報取得
#[derive(Debug, Serialize)]
pub struct StatsData {
    pub total_count: i64,
    pub in_progress_count: usize,
    pub pending_count: usize,
    pub completed_count: usize,
    pub trash_count: usize,
}

pub async fn get_stats(State(state): State<Arc<AppState>>) -> impl IntoResponse {
    match state.db.list_records(None, None, None, false) {
        Ok(records) => {
            let total = records.len();
            let in_progress = records.iter().filter(|r| r.status == "進行中").count();
            let pending = records
                .iter()
                .filter(|r| r.status == "日程未定" || r.status == "リスケ")
                .count();
            let completed = records.iter().filter(|r| r.status == "完了").count();
            let trash_count = state
                .db
                .list_deleted_records()
                .map(|v| v.len())
                .unwrap_or(0);

            (
                StatusCode::OK,
                Json(ApiResponse::ok(StatsData {
                    total_count: total as i64,
                    in_progress_count: in_progress,
                    pending_count: pending,
                    completed_count: completed,
                    trash_count,
                })),
            )
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}
