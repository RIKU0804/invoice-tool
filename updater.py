"""
自動アップデートチェッカー

GitHub Releases の最新バージョンと現在のバージョンを比較し、
新しいバージョンがある場合はダウンロードを促す。
"""

import sys
import os
import json
import logging
import urllib.request
from pathlib import Path
from typing import Callable, Optional
from version import VERSION

GITHUB_REPO = "RIKU0804/invoice-tool"
RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def _log_dir() -> Path:
    """ログ出力先: %APPDATA%\\invoice-tool (Win) or ~/.invoice-tool"""
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home())) / "invoice-tool"
    else:
        base = Path.home() / ".invoice-tool"
    base.mkdir(parents=True, exist_ok=True)
    return base


_logger = logging.getLogger("shiharai.updater")
if not _logger.handlers:
    _logger.setLevel(logging.INFO)
    try:
        _fh = logging.FileHandler(_log_dir() / "update.log", encoding="utf-8")
        _fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        _logger.addHandler(_fh)
    except Exception:
        pass  # ログハンドラ作成失敗は無視（更新チェック自体は継続）


# コールバック契約:
#   kind="new":     payload = {"version": str, "url": str}
#   kind="current": payload = str (メッセージ)
#   kind="error":   payload = str (エラー詳細)
UpdateCallback = Callable[[str, object], None]

_update_callback: Optional[UpdateCallback] = None


def set_update_callback(callback: UpdateCallback) -> None:
    global _update_callback
    _update_callback = callback


def _is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def _fetch_latest_release() -> Optional[dict]:
    try:
        req = urllib.request.Request(
            RELEASES_API,
            headers={"User-Agent": "invoice-tool", "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
            _logger.info(f"GitHub API OK: tag={data.get('tag_name')}")
            return data
    except Exception as e:
        _logger.exception(f"GitHub API fetch failed: {type(e).__name__}: {e}")
        return None


def _compare_versions(current: str, latest: str) -> bool:
    def _parse(v: str):
        return tuple(int(x) for x in v.lstrip("v").split("."))
    try:
        return _parse(latest) > _parse(current)
    except ValueError as e:
        _logger.warning(f"version parse failed: current={current} latest={latest} err={e}")
        return False


def _emit(kind: str, payload: object) -> None:
    """コールバックがあれば呼ぶ。無ければログだけ。"""
    if _update_callback:
        try:
            _update_callback(kind, payload)
        except Exception:
            _logger.exception("update_callback raised")
    else:
        _logger.info(f"no callback set; kind={kind} payload={payload}")


def get_latest_exe_asset() -> Optional[dict]:
    """最新リリースから .exe asset を取得する。"""
    release = _fetch_latest_release()
    if not release:
        return None
    for asset in release.get("assets", []):
        if asset.get("name", "").endswith(".exe"):
            return {
                "version": release.get("tag_name", ""),
                "url": asset.get("browser_download_url", ""),
                "name": asset.get("name", ""),
                "size": asset.get("size", 0),
            }
    return None


def download_exe(url: str, dest_path: Path, progress_cb: Optional[Callable[[int, int], None]] = None) -> bool:
    """
    新しいexeをダウンロードする。
    progress_cb(downloaded_bytes, total_bytes) が与えられれば進捗通知する。
    失敗時は部分ダウンロードファイルを削除する。
    """
    downloaded = 0
    total = 0
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "invoice-tool"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            with open(dest_path, "wb") as f:
                while True:
                    chunk = resp.read(8192)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_cb:
                        progress_cb(downloaded, total)
        # サイズ検証: Content-Length と実ダウンロードサイズが一致するか
        if total > 0 and downloaded != total:
            raise RuntimeError(f"サイズ不一致: {downloaded}/{total} bytes")
        _logger.info(f"downloaded {dest_path} ({downloaded} bytes)")
        return True
    except Exception:
        _logger.exception(f"exe download failed (downloaded {downloaded}/{total})")
        try:
            if dest_path.exists():
                dest_path.unlink()  # 部分ファイルを削除
        except Exception:
            pass
        return False


def perform_self_update_swap(new_exe_tmp: Path) -> None:
    """
    ダウンロード済みの新exeに差し替えて再起動する。
    - 現在のexeを新exeで上書き（batch経由、2秒待機）
    - 新exeを起動
    - このプロセスは sys.exit(0) する
    """
    import subprocess
    if not _is_frozen():
        raise RuntimeError("self-updateは.exe実行時のみ対応")
    if not new_exe_tmp.exists():
        raise RuntimeError(f"ダウンロード済みexeが見つかりません: {new_exe_tmp}")

    current_exe = Path(sys.executable).resolve()
    tmp_dir = new_exe_tmp.parent
    bat_path = tmp_dir / "invoice-tool-update.bat"

    bat_content = (
        "@echo off\r\n"
        "timeout /t 2 /nobreak >nul\r\n"
        f'move /Y "{new_exe_tmp}" "{current_exe}" >nul 2>&1\r\n'
        "if errorlevel 1 (\r\n"
        "  timeout /t 2 /nobreak >nul\r\n"
        f'  move /Y "{new_exe_tmp}" "{current_exe}" >nul\r\n'
        ")\r\n"
        f'start "" "{current_exe}"\r\n'
        'del "%~f0"\r\n'
    )
    try:
        bat_path.write_text(bat_content, encoding="cp932")
    except Exception as e:
        _logger.exception("failed to write update batch")
        raise RuntimeError(f"更新バッチの作成に失敗: {e}") from e

    _logger.info(f"launching update batch: {bat_path}")

    DETACHED_PROCESS = 0x00000008
    CREATE_NO_WINDOW = 0x08000000
    try:
        subprocess.Popen(
            ["cmd", "/c", str(bat_path)],
            creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW,
            close_fds=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        _logger.exception("failed to launch update batch")
        raise RuntimeError(f"更新バッチの起動に失敗: {e}") from e

    sys.exit(0)


def run_update_check(silent_if_current: bool = True, force: bool = False) -> None:
    """更新チェック。
    force=True: 開発環境でも実行。
    silent_if_current=False: 最新/エラー時もコールバックを呼ぶ。
    """
    _logger.info(f"run_update_check(silent={silent_if_current}, force={force}, frozen={_is_frozen()}, version=v{VERSION})")

    if not force and not _is_frozen():
        _logger.info("skip: not frozen and force=False")
        return

    release = _fetch_latest_release()
    if not release:
        if not silent_if_current:
            _emit("error", f"GitHub API に接続できませんでした。ログ: {_log_dir() / 'update.log'}")
        return

    latest_version = release.get("tag_name", "")
    if not latest_version:
        _logger.warning("release has no tag_name")
        if not silent_if_current:
            _emit("error", "最新リリース情報が取得できませんでした")
        return

    if not _compare_versions(VERSION, latest_version):
        _logger.info(f"up to date: v{VERSION} >= {latest_version}")
        if not silent_if_current:
            _emit("current", f"最新バージョンです (v{VERSION})")
        return

    html_url = release.get("html_url", f"https://github.com/{GITHUB_REPO}/releases")
    _logger.info(f"new version available: {latest_version} -> {html_url}")
    _emit("new", {"version": latest_version, "url": html_url})
