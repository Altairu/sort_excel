mod api;
mod backup;
mod db;
mod excel_exporter;
mod excel_importer;
mod models;
mod static_files;
#[cfg(windows)]
mod tray;

use anyhow::Result;
use api::AppState;
use axum::{
    routing::{delete, get, post},
    Router,
};
use chrono::{Local, Timelike};
use db::Db;
use std::net::SocketAddr;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::Duration;

#[tokio::main]
async fn main() -> Result<()> {
    // exeのあるフォルダを作業ディレクトリに確実に設定
    if let Ok(exe_path) = std::env::current_exe() {
        if let Some(exe_dir) = exe_path.parent() {
            let _ = std::env::set_current_dir(exe_dir);
        }
    }

    tracing_subscriber::fmt()
        .with_env_filter(
            tracing_subscriber::EnvFilter::try_from_default_env()
                .unwrap_or_else(|_| "info,server=debug".into()),
        )
        .init();

    tracing::info!("受託確認票 サーバー起動処理を開始...");

    let db_path = PathBuf::from("data.db");
    let backup_dir = PathBuf::from("backups");
    let exports_dir = PathBuf::from("exports");

    std::fs::create_dir_all(&backup_dir)?;
    std::fs::create_dir_all(&exports_dir)?;

    let db = Db::open(db_path.to_str().unwrap())?;

    // 初回起動時チェック: データが0件かつサンプルデータが存在する場合、自動取り込み
    if db.count_records()? == 0 {
        let sample_candidates = [
            PathBuf::from(r"Sample data\2026.10.20 受託確認票.xlsx"),
            PathBuf::from(r"..\Sample data\2026.10.20 受託確認票.xlsx"),
        ];
        for candidate in &sample_candidates {
            if candidate.exists() {
                tracing::info!(
                    "初期サンプルデータを検出しました。自動取り込みを開始します: {:?}",
                    candidate
                );
                match excel_importer::import_excel_file(candidate, &db) {
                    Ok(cnt) => tracing::info!("初回インポート完了: {} 件", cnt),
                    Err(e) => tracing::warn!("初回インポートに失敗しました: {:?}", e),
                }
                break;
            }
        }
    }

    // 起動時の初期Excelエクスポート
    if let Ok(config) = db.get_config() {
        if let Ok(records) = db.list_records(None, None, None, false) {
            let export_dir = PathBuf::from(&config.export_dir);
            let _ = excel_exporter::export_records_to_excel(
                &records,
                &config,
                &export_dir,
                &config.export_file_name,
            );
        }
    }

    let port = 3000;

    let state = Arc::new(AppState {
        db: db.clone(),
        db_path: db_path.clone(),
        backup_dir: backup_dir.clone(),
        export_lock: Arc::new(Mutex::new(())),
        port,
    });

    // Windowsタスクトレイ（インジケーター）常駐スレッドを起動
    #[cfg(windows)]
    tray::win::start_tray_loop();

    // 日次バックアップスケジューラ
    let state_backup = state.clone();
    tokio::spawn(async move {
        loop {
            tokio::time::sleep(Duration::from_secs(3600)).await;
            let now = Local::now();
            if now.hour() == 0 {
                let cfg = state_backup.db.get_config().unwrap_or_default();
                let mgr = backup::BackupManager::new(
                    &state_backup.backup_dir,
                    cfg.backup_retention_days,
                );
                let excel_path = Path::new(&cfg.export_dir).join(&cfg.export_file_name);
                if let Err(e) = mgr.create_daily_backup(&excel_path, &state_backup.db_path) {
                    tracing::error!("日次バックアップ実行エラー: {:?}", e);
                }
            }
        }
    });

    // CORSヘッダーは付与せず同一オリジン通信のみに制限（外部悪意サイトからのAPI読み出しを遮断）
    let app = Router::new()
        .route("/admin", get(static_files::admin_page_handler))
        .route("/api/session", get(api::get_session_info))
        .route(
            "/api/records",
            get(api::list_records).post(api::create_record),
        )
        .route(
            "/api/records/{id}",
            get(api::get_record)
                .put(api::update_record)
                .delete(api::delete_record),
        )
        .route("/api/config", get(api::get_config).post(api::update_config))
        .route("/api/export", post(api::trigger_export))
        .route("/api/import", post(api::import_records))
        .route(
            "/api/backups",
            get(api::list_backups).post(api::create_backup_now),
        )
        .route("/api/backups/restore", post(api::restore_backup))
        .route("/api/trash", get(api::list_trash))
        .route("/api/trash/{id}/restore", post(api::restore_trash_record))
        .route("/api/trash/{id}", delete(api::purge_trash_record))
        .route("/api/sorted-views", get(api::get_sorted_views))
        .route("/api/stats", get(api::get_stats))
        .fallback(static_files::static_handler)
        .with_state(state);

    let addr = SocketAddr::from(([0, 0, 0, 0], port));
    let listener = tokio::net::TcpListener::bind(addr).await?;

    let share_urls = api::detect_access_urls(port);

    println!("============================================================");
    println!(" 受託確認票 入力 / 自動ソートサーバーが正常に起動しました");
    println!(" 管理者画面 (このPC専用): http://localhost:{}/admin", port);
    for (idx, u) in share_urls.iter().enumerate() {
        if idx == 0 {
            println!(" 社員共有用URL (IP):      {}", u);
        } else {
            println!(" 社員共有用URL (PC名):    {}", u);
        }
    }
    println!(" 共有フォルダ出力: ./exports/受託確認票.xlsx");
    println!(" バックアップ保存: ./backups/");
    println!("============================================================");
    println!("※ この画面を最小化しておけば、バックグラウンドで動き続けます。");
    println!("※ 画面右下のタスクトレイからも使用者画面 / 管理者画面を開けます。");
    println!("※ 終了する場合は、このウィンドウを閉じるか Ctrl+C を押してください。");
    println!("============================================================");

    // 起動時に既定ブラウザで使用画面を開く
    let _ = open::that(format!("http://localhost:{}", port));

    axum::serve(
        listener,
        app.into_make_service_with_connect_info::<SocketAddr>(),
    )
    .await?;

    Ok(())
}
