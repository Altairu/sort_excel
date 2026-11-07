use crate::backup::BackupManager;
use crate::db::Db;
use crate::excel_exporter::{build_sorted_views, export_records_to_excel};
use crate::excel_importer::import_excel_file;
use crate::models::{AppConfig, Record, RecordInput};
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
    pub port: u16,
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

/// 接続元IPをわかりやすい文字列に整形
fn format_client_ip(addr: &SocketAddr) -> String {
    let ip = addr.ip();
    if ip.is_loopback() {
        format!("{} (サーバーPC)", ip)
    } else {
        ip.to_string()
    }
}

/// 更新前後の差分テキストを生成
fn build_record_diff(old: &Record, input: &RecordInput) -> String {
    let mut diffs = Vec::new();
    let mut check = |label: &str, old_val: &str, new_opt: &Option<String>| {
        if let Some(new_val) = new_opt {
            if old_val != new_val {
                diffs.push(format!("{}: {} → {}", label, if old_val.is_empty() { "(空)" } else { old_val }, if new_val.is_empty() { "(空)" } else { new_val }));
            }
        }
    };

    check("ステータス", &old.status, &input.status);
    check("担当", &old.staff, &input.staff);
    check("受諾日", &old.acceptance_date, &input.acceptance_date);
    check("形態", &old.form_type, &input.form_type);
    check("お客様氏名", &old.client_name, &input.client_name);
    check("物件", &old.property_name, &input.property_name);
    check("依頼内容", &old.request_content, &input.request_content);
    check("管轄", &old.jurisdiction, &input.jurisdiction);
    check("融資実行日", &old.loan_date, &input.loan_date);
    check("実行日補足", &old.loan_date_note, &input.loan_date_note);
    check("立会時間", &old.attendance_time, &input.attendance_time);
    check("立会場所", &old.attendance_place, &input.attendance_place);
    check("立会者", &old.attendee, &input.attendee);
    check("当日申請", &old.same_day_application, &input.same_day_application);
    check("実行時間", &old.execution_time, &input.execution_time);
    check("金消日", &old.settlement_date, &input.settlement_date);
    check("金消時間", &old.settlement_time, &input.settlement_time);
    check("金消日補足", &old.settlement_date_note, &input.settlement_date_note);
    check("金消場所", &old.settlement_place, &input.settlement_place);
    check("意思確認", &old.confirmation, &input.confirmation);
    check("書類チェック", &old.doc_check, &input.doc_check);
    check("仲介", &old.broker, &input.broker);
    check("代理店", &old.agency, &input.agency);
    check("先方書士", &old.opponent_scrivener, &input.opponent_scrivener);
    check("備考", &old.remarks, &input.remarks);
    check("申請書", &old.progress_application, &input.progress_application);
    check("必要書類", &old.progress_required_docs, &input.progress_required_docs);
    check("移転書類2", &old.progress_transfer_docs, &input.progress_transfer_docs);
    check("完了書類", &old.progress_complete_docs, &input.progress_complete_docs);
    check("受取日", &old.progress_receive_date, &input.progress_receive_date);
    check("東京へ完了", &old.progress_tokyo_complete, &input.progress_tokyo_complete);
    check("入金日", &old.progress_payment_date, &input.progress_payment_date);

    if diffs.is_empty() {
        "変更なし（上書き保存）".to_string()
    } else {
        diffs.join(" / ")
    }
}

/// サーバーPCの社内LAN向けIPアドレスとホスト名を検出してアクセスURL一覧を生成
pub fn detect_access_urls(port: u16) -> Vec<String> {
    let mut urls = Vec::new();

    // UDPソケットのルーティング先からメインのLAN IPv4アドレスを取得（外部への実通信は発生しない）
    if let Ok(socket) = std::net::UdpSocket::bind("0.0.0.0:0") {
        if socket.connect("8.8.8.8:80").is_ok() {
            if let Ok(local_addr) = socket.local_addr() {
                let ip = local_addr.ip();
                if !ip.is_loopback() && !ip.is_unspecified() {
                    urls.push(format!("http://{}:{}/", ip, port));
                }
            }
        }
    }

    // Windows環境変数 COMPUTERNAME からホスト名URLも生成
    if let Ok(comp_name) = std::env::var("COMPUTERNAME") {
        let name = comp_name.trim();
        if !name.is_empty() {
            let host_url = format!("http://{}:{}/", name, port);
            if !urls.contains(&host_url) {
                urls.push(host_url);
            }
        }
    }

    if urls.is_empty() {
        urls.push(format!("http://localhost:{}/", port));
    }

    urls
}

/// 接続元情報を取得（使用者画面で管理者リンクを表示するかどうかの判定 + 社員共有URLの取得用）
#[derive(Debug, Serialize)]
pub struct SessionInfo {
    pub is_admin_pc: bool,
    pub share_urls: Vec<String>,
}

#[derive(Debug, Deserialize)]
pub struct SessionQuery {
    pub actor: Option<String>,
    pub log_access: Option<bool>,
}

pub async fn get_session_info(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Query(q): Query<SessionQuery>,
) -> impl IntoResponse {
    if q.log_access.unwrap_or(false) {
        let actor = q.actor.as_deref().unwrap_or("");
        if !actor.is_empty() {
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                actor,
                "ACCESS",
                None,
                "画面へアクセス",
                if is_local_request(&addr) {
                    "サーバーPCから画面を開きました"
                } else {
                    "社内LANのPCから画面を開きました"
                },
            );
        }
    }

    (
        StatusCode::OK,
        Json(ApiResponse::ok(SessionInfo {
            is_admin_pc: is_local_request(&addr),
            share_urls: detect_access_urls(state.port),
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
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Json(input): Json<RecordInput>,
) -> impl IntoResponse {
    match state.db.insert_record(&input) {
        Ok(id) => {
            let actor = input.actor_name.as_deref().unwrap_or("");
            let client_name = input.client_name.as_deref().unwrap_or("(氏名未入力)");
            let staff = input.staff.as_deref().unwrap_or("-");
            let loan_date = input.loan_date.as_deref().unwrap_or("-");
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                actor,
                "CREATE",
                Some(id),
                &format!("新規登録: {} 様 (ID:{})", client_name, id),
                &format!("担当: {} / 融資実行日: {} / 物件: {}", staff, loan_date, input.property_name.as_deref().unwrap_or("-")),
            );
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
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Path(id): Path<i64>,
    Json(input): Json<RecordInput>,
) -> impl IntoResponse {
    let old_rec = state.db.get_record(id).ok().flatten();
    match state.db.update_record(id, &input) {
        Ok(()) => {
            let actor = input.actor_name.as_deref().unwrap_or("");
            let client_name = input
                .client_name
                .as_deref()
                .or_else(|| old_rec.as_ref().map(|r| r.client_name.as_str()))
                .unwrap_or("-");
            let diff_text = match &old_rec {
                Some(old) => build_record_diff(old, &input),
                None => "案件情報を更新".to_string(),
            };
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                actor,
                "UPDATE",
                Some(id),
                &format!("案件更新: {} 様 (ID:{})", client_name, id),
                &diff_text,
            );
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

#[derive(Debug, Deserialize)]
pub struct ActorQuery {
    pub actor: Option<String>,
}

/// 案件削除 (論理削除)
pub async fn delete_record(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Path(id): Path<i64>,
    Query(q): Query<ActorQuery>,
) -> impl IntoResponse {
    let old_rec = state.db.get_record(id).ok().flatten();
    match state.db.delete_record(id) {
        Ok(()) => {
            let actor = q.actor.as_deref().unwrap_or("");
            let client_name = old_rec.as_ref().map(|r| r.client_name.as_str()).unwrap_or("-");
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                actor,
                "DELETE",
                Some(id),
                &format!("ゴミ箱へ削除: {} 様 (ID:{})", client_name, id),
                "案件をゴミ箱へ移動しました",
            );
            trigger_auto_export(&state);
            (StatusCode::OK, Json(ApiResponse::ok("削除完了")))
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

#[derive(Debug, Deserialize)]
pub struct AddActorRequest {
    pub name: String,
}

/// 新しい操作者（担当者）名をリストへ追加（全クライアントから利用可）
pub async fn add_actor(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Json(payload): Json<AddActorRequest>,
) -> impl IntoResponse {
    let name = payload.name.trim().to_string();
    if name.is_empty() || name.len() > 40 {
        return (
            StatusCode::BAD_REQUEST,
            Json(ApiResponse::err("名前を正しく入力してください（40文字以内）")),
        );
    }

    match state.db.get_config() {
        Ok(mut cfg) => {
            if !cfg.staff_order.contains(&name) {
                cfg.staff_order.push(name.clone());
                if let Err(e) = state.db.update_config(&cfg) {
                    return (
                        StatusCode::INTERNAL_SERVER_ERROR,
                        Json(ApiResponse::err(e.to_string())),
                    );
                }
                let _ = state.db.insert_audit_log(
                    &format_client_ip(&addr),
                    &name,
                    "CONFIG",
                    None,
                    &format!("新しい操作者（担当者）を追加: {}", name),
                    "操作者選択画面から担当者リストへ新規追加されました",
                );
            }
            (StatusCode::OK, Json(ApiResponse::ok(cfg)))
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
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                "管理者",
                "CONFIG",
                None,
                "管理者設定を更新",
                &format!("出力先: {} / ファイル名: {} / 保持日数: {}日", cfg.export_dir, cfg.export_file_name, cfg.backup_retention_days),
            );
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
pub async fn trigger_export(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Query(q): Query<ActorQuery>,
) -> impl IntoResponse {
    let _guard = match state.export_lock.lock() {
        Ok(g) => g,
        Err(poisoned) => poisoned.into_inner(),
    };

    match state.db.get_config() {
        Ok(cfg) => match state.db.list_records(None, None, None, false) {
            Ok(records) => {
                let export_dir = PathBuf::from(&cfg.export_dir);
                match export_records_to_excel(&records, &cfg, &export_dir, &cfg.export_file_name) {
                    Ok(path) => {
                        let actor = q.actor.as_deref().unwrap_or("");
                        let _ = state.db.insert_audit_log(
                            &format_client_ip(&addr),
                            actor,
                            "EXPORT",
                            None,
                            "手動でExcel出力を実行",
                            &format!("出力先: {:?}", path),
                        );
                        (
                            StatusCode::OK,
                            Json(ApiResponse::ok(format!(
                                "Excelファイルを出力しました: {:?}",
                                path
                            ))),
                        )
                    }
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
            let mode_label = if is_replace { "総入れ替えで" } else { "追加で" };
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                "管理者",
                "IMPORT",
                None,
                &format!("Excel一括インポート ({}{})", mode_label, count),
                &format!("対象ファイル: {}", path_str),
            );
            trigger_auto_export(&state);
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
        Ok(()) => {
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                "管理者",
                "BACKUP",
                None,
                "手動バックアップを作成",
                "本日分の日次バックアップを作成しました",
            );
            (
                StatusCode::OK,
                Json(ApiResponse::ok("本日分のバックアップを作成しました")),
            )
        }
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
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                "管理者",
                "BACKUP",
                None,
                &format!("バックアップから復元 ({})", date_str),
                &format!("復元元ファイル: {:?}", backup_db_path),
            );
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
            let rec_name = state
                .db
                .get_record(id)
                .ok()
                .flatten()
                .map(|r| r.client_name)
                .unwrap_or_else(|| "-".into());
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                "管理者",
                "RESTORE",
                Some(id),
                &format!("ゴミ箱から復元: {} 様 (ID:{})", rec_name, id),
                "削除済み案件を有効状態へ戻しました",
            );
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

    let rec_name = state
        .db
        .get_record(id)
        .ok()
        .flatten()
        .map(|r| r.client_name)
        .unwrap_or_else(|| "-".into());

    match state.db.purge_record(id) {
        Ok(()) => {
            let _ = state.db.insert_audit_log(
                &format_client_ip(&addr),
                "管理者",
                "PURGE",
                Some(id),
                &format!("完全削除: {} 様 (ID:{})", rec_name, id),
                "データベースから物理削除しました",
            );
            (
                StatusCode::OK,
                Json(ApiResponse::ok("案件を完全に削除しました")),
            )
        }
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

#[derive(Debug, Deserialize)]
pub struct AuditLogQuery {
    pub action_type: Option<String>,
    pub q: Option<String>,
    pub limit: Option<usize>,
}

/// アクセス・操作履歴ログ一覧取得（管理者PCのみ許可）
pub async fn list_audit_logs(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Query(q): Query<AuditLogQuery>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err(
                "操作履歴ログはサーバーPC（管理者画面）からのみ閲覧できます",
            )),
        );
    }

    let limit = q.limit.unwrap_or(500);
    match state
        .db
        .list_audit_logs(q.action_type.as_deref(), q.q.as_deref(), limit)
    {
        Ok(logs) => (StatusCode::OK, Json(ApiResponse::ok(logs))),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// 監査ログの単一削除（管理者PCのみ許可）
pub async fn delete_audit_log_entry(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Path(id): Path<i64>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err("管理者画面からのみ実行できます")),
        );
    }

    match state.db.delete_audit_log(id) {
        Ok(()) => (StatusCode::OK, Json(ApiResponse::ok("ログを削除しました"))),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

#[derive(Debug, Deserialize)]
pub struct ClearLogsRequest {
    pub only_test: Option<bool>,
}

/// 監査ログの一括消去（テストログのみ、または全件 / 管理者PCのみ許可）
pub async fn clear_audit_logs_handler(
    ConnectInfo(addr): ConnectInfo<SocketAddr>,
    State(state): State<Arc<AppState>>,
    Json(payload): Json<ClearLogsRequest>,
) -> impl IntoResponse {
    if !is_local_request(&addr) {
        return (
            StatusCode::FORBIDDEN,
            Json(ApiResponse::err("管理者画面からのみ実行できます")),
        );
    }

    let only_test = payload.only_test.unwrap_or(true);
    match state.db.clear_audit_logs(only_test) {
        Ok(cnt) => (
            StatusCode::OK,
            Json(ApiResponse::ok(format!("{} 件のログを消去しました", cnt))),
        ),
        Err(e) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(ApiResponse::err(e.to_string())),
        ),
    }
}

/// ソート済みデータ一式取得（日付別シート / 融資実行日順 / 金消日順 / 統合データ）
pub async fn get_sorted_views(State(state): State<Arc<AppState>>) -> impl IntoResponse {
    match state.db.get_config() {
        Ok(cfg) => match state.db.list_records(None, None, None, false) {
            Ok(records) => {
                let views = build_sorted_views(&records, &cfg);
                (StatusCode::OK, Json(ApiResponse::ok(views)))
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
