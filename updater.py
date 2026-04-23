"""
自動アップデートチェッカー

GitHub Releases の最新バージョンと現在のバージョンを比較し、
新しいバージョンがある場合はダウンロードを促す。
PyInstaller でビルドされた .exe の場合のみ動作する。
"""

import sys
import os
import json
import urllib.request
from version import VERSION

GITHUB_REPO = "RIKU0804/invoice-tool"
RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

_update_callback = None  # GUI側からセットされるコールバック(version, url) -> None


def set_update_callback(callback):
    global _update_callback
    _update_callback = callback


def _is_frozen() -> bool:
    return getattr(sys, "frozen", False)


_last_error: str | None = None


def get_last_error() -> str | None:
    return _last_error


def _fetch_latest_release() -> dict | None:
    global _last_error
    try:
        req = urllib.request.Request(
            RELEASES_API,
            headers={"User-Agent": "shiharai-tool", "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            _last_error = None
            return json.loads(resp.read())
    except Exception as e:
        _last_error = f"{type(e).__name__}: {e}"
        return None


def _compare_versions(current: str, latest: str) -> bool:
    def _parse(v: str):
        return tuple(int(x) for x in v.lstrip("v").split("."))
    try:
        return _parse(latest) > _parse(current)
    except ValueError:
        return False


def run_update_check(silent_if_current: bool = True, force: bool = False) -> None:
    # force=Trueなら開発環境でも実行
    if not force and not _is_frozen():
        return

    release = _fetch_latest_release()
    if not release:
        if _update_callback and not silent_if_current:
            _update_callback("error", f"更新確認に失敗: {_last_error or 'unknown'}")
        return

    latest_version = release.get("tag_name", "")
    if not latest_version:
        return

    if not _compare_versions(VERSION, latest_version):
        if not silent_if_current:
            if _update_callback:
                _update_callback("current", f"最新バージョンです (v{VERSION})")
            else:
                print(f"最新バージョンです (v{VERSION})")
        return

    # 新バージョンあり
    html_url = release.get("html_url", f"https://github.com/{GITHUB_REPO}/releases")

    # コールバックがあればGUI側に委譲（customtkinter対応）
    if _update_callback:
        _update_callback(latest_version, html_url)
        return

    # フォールバック: tkinter
    try:
        import tkinter as tk
        from tkinter import messagebox
        import webbrowser

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        answer = messagebox.askyesno(
            "アップデートのお知らせ",
            f"新しいバージョン {latest_version} が利用可能です。\n"
            f"現在のバージョン: v{VERSION}\n\n"
            "ダウンロードページを開きますか？",
            parent=root,
        )
        root.destroy()
        if answer:
            webbrowser.open(html_url)
    except Exception:
        print(f"新バージョン利用可能: {latest_version}")
