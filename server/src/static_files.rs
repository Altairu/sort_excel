use axum::{
    http::{header, StatusCode, Uri},
    response::{IntoResponse, Response},
};
use rust_embed::Embed;

#[derive(Embed)]
#[folder = "static/"]
struct Assets;

pub async fn static_handler(uri: Uri) -> Response {
    let mut path = uri.path().trim_start_matches('/').to_string();
    if path.is_empty() {
        path = "index.html".to_string();
    }

    match Assets::get(&path) {
        Some(content) => {
            let mime = mime_guess::from_path(&path).first_or_octet_stream();
            ([(header::CONTENT_TYPE, mime.as_ref())], content.data).into_response()
        }
        None => {
            // SPAフォールバック
            if let Some(content) = Assets::get("index.html") {
                ([(header::CONTENT_TYPE, "text/html; charset=utf-8")], content.data).into_response()
            } else {
                (StatusCode::NOT_FOUND, "Not Found").into_response()
            }
        }
    }
}
