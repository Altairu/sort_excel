use crate::backup::BackupManager;
use crate::db::Db;
use crate::excel_exporter::export_records_to_excel;
use crate::excel_importer::import_excel_file;
use crate::models::{AppConfig, RecordInput};
use axum::{
    extract::{Path, Query, State},
    http::StatusCode,
    response::IntoResponse,
    Json,
};
use serde::{Deserialize, Serialize};
use std::path::{Path as StdPath, PathBuf};
use std::sync::Arc;

#[derive(Clone)]
pub struct AppState {
    pub db: Db,
    pub db_path: PathBuf,
    pub backup_dir: PathBuf,
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

/// 自動エクスポートヘルパー
fn trigger_auto_export(state: &AppState) {
    let db = state.db.clone();

    tokio::spawn(async move {
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

/// 設定取得
pub async fn get_config(State(state): State<Arc<AppState>>) -> impl IntoResponse {
    match state.db.get_config() {
        Ok(cfg) => (StatusCode::OK, Json(ApiResponse::ok(cfg))),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 設定更新
pub async fn update_config(
    State(state): State<Arc<AppState>>,
    Json(cfg): Json<AppConfig>,
) -> impl IntoResponse {
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
}

/// 既存Excelファイルのインポート (パス指定またはマルチパート)
pub async fn import_records(
    State(state): State<Arc<AppState>>,
    Json(payload): Json<ImportRequest>,
) -> impl IntoResponse {
    let path_str = payload
        .file_path
        .unwrap_or_else(|| r"Sample data\2026.10.20 受託確認票.xlsx".to_string());
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

    match import_excel_file(path, &state.db) {
        Ok(count) => {
            trigger_auto_export(&state);
            (
                StatusCode::OK,
                Json(ApiResponse::ok(format!(
                    "{} 件の案件を取り込みました",
                    count
                ))),
            )
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(format!("インポートエラー: {}", e))),
        ),
    }
}

/// バックアップ一覧取得
pub async fn list_backups(State(state): State<Arc<AppState>>) -> impl IntoResponse {
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

/// 統計情報取得
#[derive(Debug, Serialize)]
pub struct StatsData {
    pub total_count: i64,
    pub in_progress_count: usize,
    pub pending_count: usize,
    pub completed_count: usize,
}

pub async fn get_stats(State(state): State<Arc<AppState>>) -> impl IntoResponse {
    match state.db.list_records(None, None, None, false) {
        Ok(records) => {
            let total = records.len();
            let in_progress = records.iter().filter(|r| r.status == "進行中").count();
            let pending = records.iter().filter(|r| r.status == "日程未定" || r.status == "リスケ").count();
            let completed = records.iter().filter(|r| r.status == "完了").count();

            (
                StatusCode::OK,
                Json(ApiResponse::ok(StatsData {
                    total_count: total as i64,
                    in_progress_count: in_progress,
                    pending_count: pending,
                    completed_count: completed,
                })),
            )
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}
