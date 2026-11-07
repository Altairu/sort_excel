mod api;
mod backup;
mod db;
mod excel_exporter;
mod excel_importer;
mod models;
mod static_files;

use anyhow::Result;
use api::AppState;
use axum::{
    routing::{get, post},
    Router,
};
use chrono::{Local, Timelike};
use db::Db;
use std::net::SocketAddr;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Duration;
use tower_http::cors::CorsLayer;

#[tokio::main]
async fn main() -> Result<()> {
    tracing_subscriber::fmt()
        .with_env_filter(
            tracing_subscriber::EnvFilter::try_from_default_env()
                .unwrap_or_else(|_| "info,server=debug".into()),
        )
        .init();

    tracing::info!("受託確認票 Webサーバーを起動中...");

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
            PathBuf::from(r"C:\Users\106no\Documents\GitHub\sort_excel\Sample data\2026.10.20 受託確認票.xlsx"),
        ];
        for candidate in &sample_candidates {
            if candidate.exists() {
                tracing::info!("初期サンプルデータを検出しました。自動取り込みを開始します: {:?}", candidate);
                match excel_importer::import_excel_file(candidate, &db) {
                    Ok(cnt) => tracing::info!("初回インポート完了: {} 件", cnt),
                    Err(e) => tracing::warn!("初回インポートに失敗しました: {:?}", e),
                }
                break;
            }
        }
    }

    let state = Arc::new(AppState {
        db: db.clone(),
        db_path: db_path.clone(),
        backup_dir: backup_dir.clone(),
    });

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

    // 日次バックアップスケジューラ (Tokio background task)
    let state_backup = state.clone();
    tokio::spawn(async move {
        loop {
            // 1時間おきにチェックし、日付変更直後（0時台）に日次バックアップを実行
            tokio::time::sleep(Duration::from_secs(3600)).await;
            let now = Local::now();
            if now.hour() == 0 {
                let cfg = state_backup.db.get_config().unwrap_or_default();
                let mgr = backup::BackupManager::new(&state_backup.backup_dir, cfg.backup_retention_days);
                let excel_path = Path::new(&cfg.export_dir).join(&cfg.export_file_name);
                if let Err(e) = mgr.create_daily_backup(&excel_path, &state_backup.db_path) {
                    tracing::error!("日次バックアップ実行エラー: {:?}", e);
                }
            }
        }
    });

    let app = Router::new()
        .route("/api/records", get(api::list_records).post(api::create_record))
        .route(
            "/api/records/{id}",
            get(api::get_record)
                .put(api::update_record)
                .delete(api::delete_record),
        )
        .route(
            "/api/config",
            get(api::get_config).post(api::update_config),
        )
        .route("/api/export", post(api::trigger_export))
        .route("/api/import", post(api::import_records))
        .route("/api/backups", get(api::list_backups))
        .route("/api/stats", get(api::get_stats))
        .fallback(static_files::static_handler)
        .layer(CorsLayer::permissive())
        .with_state(state);

    let port = 8080;
    let addr = SocketAddr::from(([0, 0, 0, 0], port));

    println!("============================================================");
    println!(" 受託確認票 入力・自動ソートサーバーが正常に起動しました！");
    println!(" ローカルアクセス: http://localhost:{}", port);
    println!(" 社内LANアクセス:  http://[質問者様PC名]:{} または IPアドレス", port);
    println!(" 共有フォルダ出力: ./exports/受託確認票.xlsx");
    println!(" バックアップ保存: ./backups/");
    println!("============================================================");

    let listener = tokio::net::TcpListener::bind(addr).await?;
    axum::serve(listener, app).await?;

    Ok(())
}
