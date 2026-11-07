use axum::{
    extract::ConnectInfo,
    http::{header, StatusCode, Uri},
    response::{IntoResponse, Response},
};
use rust_embed::Embed;
use std::net::SocketAddr;

#[derive(Embed)]
#[folder = "static/"]
struct Assets;

pub async fn admin_page_handler(ConnectInfo(addr): ConnectInfo<SocketAddr>) -> Response {
    if !addr.ip().is_loopback() {
        return (
            StatusCode::FORBIDDEN,
            [(header::CONTENT_TYPE, "text/plain; charset=utf-8")],
            "403 Forbidden: 管理者画面はサーバーを実行しているPC（http://localhost:8080/admin）からのみアクセスできます。",
        )
            .into_response();
    }

    match Assets::get("admin.html") {
        Some(content) => (
            [(header::CONTENT_TYPE, "text/html; charset=utf-8")],
            content.data,
        )
            .into_response(),
        None => (StatusCode::NOT_FOUND, "Admin Page Not Found").into_response(),
    }
}

pub async fn static_handler(ConnectInfo(addr): ConnectInfo<SocketAddr>, uri: Uri) -> Response {
    let mut path = uri.path().trim_start_matches('/').to_string();
    if path.is_empty() {
        path = "index.html".to_string();
    }

    // admin.html や admin.js への外部からの直接アクセスを遮断
    if (path == "admin.html" || path == "admin.js" || path == "admin") && !addr.ip().is_loopback() {
        return (
            StatusCode::FORBIDDEN,
            [(header::CONTENT_TYPE, "text/plain; charset=utf-8")],
            "403 Forbidden: 管理者専用リソースです。",
        )
            .into_response();
    }

    match Assets::get(&path) {
        Some(content) => {
            let mime = mime_guess::from_path(&path).first_or_octet_stream();
            ([(header::CONTENT_TYPE, mime.as_ref())], content.data).into_response()
        }
        None => {
            // SPAフォールバック
            if let Some(content) = Assets::get("index.html") {
                (
                    [(header::CONTENT_TYPE, "text/html; charset=utf-8")],
                    content.data,
                )
                    .into_response()
            } else {
                (StatusCode::NOT_FOUND, "Not Found").into_response()
            }
        }
    }
}
