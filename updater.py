# -*- coding: utf-8 -*-
"""
自動アップデートモジュール
GitHub Releases の最新版を確認し、exe のダウンロードと差し替えを行います。
標準ライブラリのみで実装しているため、追加の依存はありません。
"""

import hashlib
import json
import os
import re
import subprocess
import sys
import time
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional, Tuple

from version import __version__

# 更新元のリポジトリとアセット名
GITHUB_REPO = "Altairu/sort_excel"
ASSET_NAME = "sort_excel.exe"
CHECKSUM_ASSET_NAME = ASSET_NAME + ".sha256"

LATEST_RELEASE_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
RELEASES_PAGE_URL = f"https://github.com/{GITHUB_REPO}/releases/latest"

USER_AGENT = f"sort_excel-updater/{__version__}"

# 差し替え時に使う一時ファイルの拡張子
NEW_SUFFIX = ".new"
OLD_SUFFIX = ".old"


class UpdateError(Exception):
    """アップデート処理中のエラー"""


@dataclass
class ReleaseInfo:
    """最新リリースの情報"""
    tag: str
    version: Tuple[int, ...]
    notes: str
    html_url: str
    exe_url: Optional[str]
    checksum_url: Optional[str]


def parse_version(text: str) -> Tuple[int, ...]:
    """v2.1.0 や V2.0.0 のような文字列を数値のタプルに変換します"""
    text = text.strip().lstrip("vV")
    parts = []
    for piece in text.split("."):
        m = re.match(r"\d+", piece)
        if not m:
            break
        parts.append(int(m.group()))
    return tuple(parts)


def current_version() -> Tuple[int, ...]:
    """実行中のアプリのバージョンを返します"""
    return parse_version(__version__)


def is_frozen() -> bool:
    """PyInstaller で作った exe として動いているかを返します"""
    return bool(getattr(sys, "frozen", False))


def _request(url: str, accept: str) -> urllib.request.Request:
    return urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": accept})


def fetch_latest_release(timeout: float = 5.0) -> ReleaseInfo:
    """GitHub API から最新の正式リリースを取得します"""
    req = _request(LATEST_RELEASE_API, "application/vnd.github+json")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as res:
            data = json.loads(res.read().decode("utf-8"))
    except Exception as e:
        raise UpdateError(f"最新版の情報を取得できませんでした: {e}") from e

    tag = data.get("tag_name", "")
    exe_url = None
    checksum_url = None
    for asset in data.get("assets", []):
        name = asset.get("name")
        if name == ASSET_NAME:
            exe_url = asset.get("browser_download_url")
        elif name == CHECKSUM_ASSET_NAME:
            checksum_url = asset.get("browser_download_url")

    return ReleaseInfo(
        tag=tag,
        version=parse_version(tag),
        notes=data.get("body") or "",
        html_url=data.get("html_url") or RELEASES_PAGE_URL,
        exe_url=exe_url,
        checksum_url=checksum_url,
    )


def is_newer(info: ReleaseInfo) -> bool:
    """取得したリリースが実行中の版より新しいかを返します"""
    return bool(info.version) and info.version > current_version()


def can_self_update(info: ReleaseInfo) -> bool:
    """このリリースで exe の自動差し替えができるかを返します"""
    return is_frozen() and bool(info.exe_url) and bool(info.checksum_url)


def _download_text(url: str, timeout: float) -> str:
    with urllib.request.urlopen(_request(url, "application/octet-stream"), timeout=timeout) as res:
        return res.read().decode("utf-8", errors="ignore")


def download_update(
    info: ReleaseInfo,
    progress_cb: Optional[Callable[[int, int], None]] = None,
    timeout: float = 30.0,
) -> Path:
    """新しい exe を実行中の exe と同じフォルダへダウンロードし、SHA256 を照合します"""
    if not can_self_update(info):
        raise UpdateError("このリリースには自動更新に必要なファイルがありません。")

    # 期待するハッシュ値を取得します。形式は "ハッシュ  ファイル名" です
    checksum_text = _download_text(info.checksum_url, timeout)
    expected = checksum_text.strip().split()[0].lower() if checksum_text.strip() else ""
    if not re.fullmatch(r"[0-9a-f]{64}", expected):
        raise UpdateError("チェックサムファイルの形式が正しくありません。")

    exe_path = Path(sys.executable).resolve()
    new_path = exe_path.with_name(exe_path.name + NEW_SUFFIX)

    sha = hashlib.sha256()
    try:
        with urllib.request.urlopen(_request(info.exe_url, "application/octet-stream"), timeout=timeout) as res:
            total = int(res.headers.get("Content-Length") or 0)
            done = 0
            with open(new_path, "wb") as f:
                while True:
                    chunk = res.read(1024 * 256)
                    if not chunk:
                        break
                    f.write(chunk)
                    sha.update(chunk)
                    done += len(chunk)
                    if progress_cb:
                        progress_cb(done, total)
    except Exception as e:
        _safe_unlink(new_path)
        raise UpdateError(f"ダウンロードに失敗しました: {e}") from e

    if sha.hexdigest().lower() != expected:
        _safe_unlink(new_path)
        raise UpdateError("ダウンロードしたファイルのハッシュが一致しません。更新を中止しました。")

    return new_path


def apply_update(new_path: Path) -> None:
    """実行中の exe を新しい exe に差し替え、新しい exe を起動します

    呼び出し元はこの関数が成功したらすぐにアプリを終了してください。
    """
    exe_path = Path(sys.executable).resolve()
    old_path = exe_path.with_name(exe_path.name + OLD_SUFFIX)

    # 前回の .old が残っていれば先に消しておきます
    _safe_unlink(old_path)

    try:
        # 実行中の exe は上書きできないが、リネームはできます
        os.replace(exe_path, old_path)
    except OSError as e:
        raise UpdateError(f"実行中のファイルを退避できませんでした: {e}") from e

    try:
        os.replace(new_path, exe_path)
    except OSError as e:
        # 失敗したら元に戻して古い版のまま使えるようにします
        try:
            os.replace(old_path, exe_path)
        except OSError:
            pass
        raise UpdateError(f"新しいファイルを配置できませんでした: {e}") from e

    # onefile の exe から自分を起動し直すときは、親の一時フォルダの情報を引き継がないようにします
    env = dict(os.environ)
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    try:
        subprocess.Popen([str(exe_path)], env=env, cwd=str(exe_path.parent), close_fds=True)
    except OSError as e:
        raise UpdateError(f"新しいバージョンを起動できませんでした。手動で起動してください: {e}") from e


def cleanup_old_files() -> None:
    """前回の更新で残った .old と .new のファイルを削除します"""
    if not is_frozen():
        return
    exe_path = Path(sys.executable).resolve()
    for suffix in (OLD_SUFFIX, NEW_SUFFIX):
        target = exe_path.with_name(exe_path.name + suffix)
        if not target.exists():
            continue
        # 古いプロセスの終了を少し待つ必要があるため、数回だけ試します
        for _ in range(5):
            if _safe_unlink(target):
                break
            time.sleep(0.5)


def _safe_unlink(path: Path) -> bool:
    try:
        if path.exists():
            path.unlink()
        return True
    except OSError:
        return False
