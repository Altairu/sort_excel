fn main() {
    if std::env::var("TARGET").map(|t| t.contains("windows")).unwrap_or(false) {
        let mut res = winres::WindowsResource::new();
        res.set_icon("../icon.ico");
        res.set("ProductName", "受託確認票 業務管理システム");
        res.set("FileDescription", "受託確認票 入力・自動ソート・バックアップ サーバー");
        res.set("LegalCopyright", "Copyright (C) 2026");
        let _ = res.compile();
    }
}
