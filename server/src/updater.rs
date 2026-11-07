use anyhow::{anyhow, Result};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::thread;
use std::time::Duration;

#[cfg(windows)]
use std::os::windows::process::CommandExt;

pub const CURRENT_VERSION: &str = env!("CARGO_PKG_VERSION");
pub const GITHUB_REPO: &str = "Altairu/sort_excel";
pub const ASSET_NAME: &str = "sort_excel_server.exe";
pub const CHECKSUM_ASSET_NAME: &str = "sort_excel_server.exe.sha256";
pub const LATEST_RELEASE_API: &str = "https://api.github.com/repos/Altairu/sort_excel/releases/latest";
pub const RELEASES_PAGE_URL: &str = "https://github.com/Altairu/sort_excel/releases/latest";

const NEW_SUFFIX: &str = ".new";
const OLD_SUFFIX: &str = ".old";

#[cfg(windows)]
const CREATE_NO_WINDOW: u32 = 0x08000000;

/// アップデート確認結果
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UpdateCheckResult {
    pub current_version: String,
    pub latest_tag: String,
    pub latest_version: String,
    pub is_update_available: bool,
    pub can_auto_update: bool,
    pub release_notes: String,
    pub release_url: String,
    pub published_at: String,
    pub exe_download_url: Option<String>,
    pub checksum_download_url: Option<String>,
}

#[derive(Debug, Deserialize)]
struct GithubRelease {
    tag_name: Option<String>,
    body: Option<String>,
    html_url: Option<String>,
    published_at: Option<String>,
    assets: Option<Vec<GithubAsset>>,
    message: Option<String>,
}

#[derive(Debug, Deserialize)]
struct GithubAsset {
    name: String,
    browser_download_url: String,
}

/// v2.2.1 や V2.0.0 のような文字列を数値のベクタに変換
pub fn parse_version(text: &str) -> Vec<u32> {
    let trimmed = text.trim().trim_start_matches(['v', 'V']);
    let mut parts = Vec::new();
    for piece in trimmed.split('.') {
        let digits: String = piece.chars().take_while(|c| c.is_ascii_digit()).collect();
        if digits.is_empty() {
            break;
        }
        if let Ok(n) = digits.parse::<u32>() {
            parts.push(n);
        } else {
            break;
        }
    }
    parts
}

/// 実行中のexeパスにサフィックスを付与したパスを返す
fn sibling_with_suffix(exe_path: &Path, suffix: &str) -> PathBuf {
    let file_name = exe_path
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| ASSET_NAME.to_string());
    exe_path.with_file_name(format!("{}{}", file_name, suffix))
}

/// 前回の自己アップデートで残った .old や .new ファイルを起動時に削除
pub fn cleanup_old_files() {
    let Ok(exe_path) = std::env::current_exe() else {
        return;
    };
    for suffix in [OLD_SUFFIX, NEW_SUFFIX] {
        let target = sibling_with_suffix(&exe_path, suffix);
        if !target.exists() {
            continue;
        }
        for _ in 0..6 {
            if fs::remove_file(&target).is_ok() || !target.exists() {
                break;
            }
            thread::sleep(Duration::from_millis(500));
        }
    }
}

/// curl.exe を使ってテキストを取得
fn fetch_url_text(url: &str, accept: &str) -> Result<String> {
    let user_agent = format!("sort_excel_server-updater/{}", CURRENT_VERSION);
    let mut cmd = Command::new("curl.exe");
    cmd.args([
        "-sSL",
        "--max-time",
        "15",
        "-H",
        &format!("User-Agent: {}", user_agent),
        "-H",
        &format!("Accept: {}", accept),
        url,
    ]);
    #[cfg(windows)]
    cmd.creation_flags(CREATE_NO_WINDOW);

    let output = cmd
        .output()
        .map_err(|e| anyhow!("通信コマンド(curl.exe)の起動に失敗しました: {}", e))?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(anyhow!("GitHubとの通信に失敗しました: {}", stderr.trim()));
    }

    Ok(String::from_utf8_lossy(&output.stdout).to_string())
}

/// curl.exe を使ってファイルを指定パスへダウンロード
fn download_url_to_file(url: &str, dest: &Path) -> Result<()> {
    let user_agent = format!("sort_excel_server-updater/{}", CURRENT_VERSION);
    let mut cmd = Command::new("curl.exe");
    cmd.args([
        "-sSL",
        "--fail",
        "--max-time",
        "120",
        "-H",
        &format!("User-Agent: {}", user_agent),
        "-H",
        "Accept: application/octet-stream",
        "-o",
        &dest.to_string_lossy(),
        url,
    ]);
    #[cfg(windows)]
    cmd.creation_flags(CREATE_NO_WINDOW);

    let output = cmd
        .output()
        .map_err(|e| anyhow!("ダウンロードプロセスの起動に失敗しました: {}", e))?;

    if !output.status.success() {
        let _ = fs::remove_file(dest);
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(anyhow!(
            "最新ファイルのダウンロードに失敗しました: {}",
            stderr.trim()
        ));
    }

    Ok(())
}

/// ファイルのSHA256ハッシュ値（小文字16進数64文字）を計算
fn compute_file_sha256(path: &Path) -> Result<String> {
    let mut file = fs::File::open(path)?;
    let mut hasher = Sha256::new();
    let mut buf = [0u8; 65536];
    loop {
        let n = file.read(&mut buf)?;
        if n == 0 {
            break;
        }
        hasher.update(&buf[..n]);
    }
    Ok(format!("{:x}", hasher.finalize()))
}

/// GitHub Releases API から最新リリース情報を確認
pub fn check_for_update() -> Result<UpdateCheckResult> {
    let body_text = fetch_url_text(LATEST_RELEASE_API, "application/vnd.github+json")?;
    let rel: GithubRelease = serde_json::from_str(&body_text)
        .map_err(|e| anyhow!("GitHub API応答の解析に失敗しました: {}", e))?;

    // リリースがまだ1つもない場合 (Not Found)
    if rel.tag_name.is_none() {
        let msg = rel.message.unwrap_or_default();
        return Ok(UpdateCheckResult {
            current_version: CURRENT_VERSION.to_string(),
            latest_tag: format!("v{}", CURRENT_VERSION),
            latest_version: CURRENT_VERSION.to_string(),
            is_update_available: false,
            can_auto_update: false,
            release_notes: if msg.contains("Not Found") {
                "現在公開されている新しいリリースはありません。".to_string()
            } else {
                format!("GitHub応答: {}", msg)
            },
            release_url: RELEASES_PAGE_URL.to_string(),
            published_at: String::new(),
            exe_download_url: None,
            checksum_download_url: None,
        });
    }

    let tag = rel.tag_name.unwrap_or_default();
    let current_ver = parse_version(CURRENT_VERSION);
    let latest_ver = parse_version(&tag);
    let is_newer = !latest_ver.is_empty() && latest_ver > current_ver;

    let mut exe_url = None;
    let mut checksum_url = None;

    if let Some(assets) = rel.assets {
        for asset in assets {
            if asset.name == ASSET_NAME {
                exe_url = Some(asset.browser_download_url);
            } else if asset.name == CHECKSUM_ASSET_NAME {
                checksum_url = Some(asset.browser_download_url);
            }
        }
    }

    let can_auto_update = exe_url.is_some() && checksum_url.is_some();

    Ok(UpdateCheckResult {
        current_version: CURRENT_VERSION.to_string(),
        latest_tag: tag.clone(),
        latest_version: tag.trim_start_matches(['v', 'V']).to_string(),
        is_update_available: is_newer,
        can_auto_update,
        release_notes: rel.body.unwrap_or_default(),
        release_url: rel.html_url.unwrap_or_else(|| RELEASES_PAGE_URL.to_string()),
        published_at: rel.published_at.unwrap_or_default(),
        exe_download_url: exe_url,
        checksum_download_url: checksum_url,
    })
}

/// 最新版のexeをダウンロードしてSHA256を検証し、実行ファイルを差し替えて再起動をスケジュールする
pub fn download_and_stage_update() -> Result<UpdateCheckResult> {
    let info = check_for_update()?;
    if !info.is_update_available {
        return Err(anyhow!(
            "現在すでに最新バージョン (v{}) を使用しています。",
            CURRENT_VERSION
        ));
    }

    let exe_url = info.exe_download_url.as_ref().ok_or_else(|| {
        anyhow!(
            "最新リリース ({}) にサーバー用ファイル ({}) が添付されていません。",
            info.latest_tag,
            ASSET_NAME
        )
    })?;
    let checksum_url = info.checksum_download_url.as_ref().ok_or_else(|| {
        anyhow!(
            "最新リリース ({}) にSHA256チェックサムファイル ({}) が添付されていません。",
            info.latest_tag,
            CHECKSUM_ASSET_NAME
        )
    })?;

    // 1. 期待されるSHA256ハッシュを取得
    let checksum_text = fetch_url_text(checksum_url, "application/octet-stream")?;
    let expected_hash = checksum_text
        .split_whitespace()
        .next()
        .unwrap_or("")
        .trim()
        .to_lowercase();

    if expected_hash.len() != 64 || !expected_hash.chars().all(|c| c.is_ascii_hexdigit()) {
        return Err(anyhow!("チェックサムファイルの形式が不正です。"));
    }

    // 2. 新しいexeを .new としてダウンロード
    let exe_path = std::env::current_exe()
        .map_err(|e| anyhow!("実行中ファイルのパス取得に失敗しました: {}", e))?;
    let new_path = sibling_with_suffix(&exe_path, NEW_SUFFIX);
    let old_path = sibling_with_suffix(&exe_path, OLD_SUFFIX);

    let _ = fs::remove_file(&new_path);
    download_url_to_file(exe_url, &new_path)?;

    // 3. SHA256ハッシュを照合
    let actual_hash = compute_file_sha256(&new_path)?;
    if actual_hash.to_lowercase() != expected_hash {
        let _ = fs::remove_file(&new_path);
        return Err(anyhow!(
            "ダウンロードしたファイルのSHA256ハッシュが一致しません（改ざんまたは破損の可能性があるため更新を中止しました）。"
        ));
    }

    // 4. 実行中exeを .old に退避し、.new を本番ファイル名へ置き換え
    let _ = fs::remove_file(&old_path);
    fs::rename(&exe_path, &old_path).map_err(|e| {
        let _ = fs::remove_file(&new_path);
        anyhow!(
            "実行中のファイルを退避できませんでした（フォルダの書き込み権限を確認してください）: {}",
            e
        )
    })?;

    if let Err(e) = fs::rename(&new_path, &exe_path) {
        let _ = fs::rename(&old_path, &exe_path);
        return Err(anyhow!("新しいファイルの配置に失敗しました: {}", e));
    }

    Ok(info)
}

/// 2秒後に新しいexeを起動し、現在のサーバープロセスを終了する
pub fn trigger_restart_into_new_binary() {
    let Ok(exe_path) = std::env::current_exe() else {
        return;
    };
    let parent_dir = exe_path
        .parent()
        .map(|p| p.to_path_buf())
        .unwrap_or_else(|| PathBuf::from("."));

    // 現在のプロセスが終了してポート3000を解放してから新しいexeを起動する
    let exe_str = exe_path.to_string_lossy().to_string();
    let mut cmd = Command::new("cmd.exe");
    cmd.args([
        "/C",
        &format!("ping 127.0.0.1 -n 3 >NUL & start \"\" \"{}\"", exe_str),
    ])
    .current_dir(parent_dir);

    #[cfg(windows)]
    cmd.creation_flags(CREATE_NO_WINDOW);

    let _ = cmd.spawn();

    thread::spawn(|| {
        thread::sleep(Duration::from_millis(500));
        std::process::exit(0);
    });
}
