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


def _fetch_latest_release() -> dict | None:
    try:
        req = urllib.request.Request(
            RELEASES_API,
            headers={"User-Agent": "shiharai-tool", "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            return json.loads(resp.read())
    except Exception:
        return None


def _compare_versions(current: str, latest: str) -> bool:
    def _parse(v: str):
        return tuple(int(x) for x in v.lstrip("v").split("."))
    try:
        return _parse(latest) > _parse(current)
    except ValueError:
        return False


def run_update_check(silent_if_current: bool = True) -> None:
    if not _is_frozen():
        return  # 開発環境では無視

    release = _fetch_latest_release()
    if not release:
        return

    latest_version = release.get("tag_name", "")
    if not latest_version:
        return

    if not _compare_versions(VERSION, latest_version):
        if not silent_if_current:
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
