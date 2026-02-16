"""OneDrive 共有リンク経由で xlsm ファイルをダウンロードする（読み取り専用）

共有リンクに &download=1 を付与し、requests.Session でリダイレクトを
追跡することで認証なしで直接ダウンロードする。
"""

import io
import re
import time

import requests


# --- キャッシュ ---
_cache = {"data": None, "fetched_at": 0.0, "filename": None, "content_length": None}
CACHE_TTL = 300  # 5 分


def _is_cache_fresh() -> bool:
    if _cache["data"] is None:
        return False
    return (time.time() - _cache["fetched_at"]) < CACHE_TTL


def _build_download_url(share_url: str) -> str:
    """共有リンクに download=1 パラメータを付与する"""
    sep = "&" if "?" in share_url else "?"
    return share_url + sep + "download=1"


def invalidate_cache():
    """キャッシュを強制クリアする（次回 fetch 時に再ダウンロード）"""
    _cache.update(data=None, fetched_at=0.0)


def fetch_xlsm(share_url: str) -> io.BytesIO:
    """共有リンクから xlsm をダウンロードし BytesIO で返す（キャッシュ付き）

    キャッシュが新鮮（5分以内）ならそのまま返す。
    """
    if _is_cache_fresh():
        _cache["data"].seek(0)
        return _cache["data"]

    # セッションを使ってリダイレクトチェーンを追跡（Cookie が必要）
    session = requests.Session()
    download_url = _build_download_url(share_url)
    resp = session.get(download_url, timeout=60)
    resp.raise_for_status()

    # レスポンスヘッダーからファイル名を取得
    filename = None
    cd = resp.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*=UTF-8''(.+?)(?:;|$)", cd)
    if m:
        from urllib.parse import unquote
        filename = unquote(m.group(1))
    else:
        m = re.search(r'filename="?([^";]+)"?', cd)
        if m:
            filename = m.group(1)

    content_type = resp.headers.get("Content-Type", "")
    final_url = resp.url

    buf = io.BytesIO(resp.content)
    _cache.update(
        data=buf, fetched_at=time.time(), filename=filename,
        content_length=len(resp.content), content_type=content_type,
        final_url=final_url, status_code=resp.status_code,
    )
    buf.seek(0)
    return buf


def get_file_info() -> dict:
    """キャッシュされたファイルの情報を返す"""
    if _cache["data"] is None:
        return {"filename": None, "fetched_at": None, "content_length": None}
    return {
        "filename": _cache.get("filename"),
        "fetched_at": _cache.get("fetched_at"),
        "content_length": _cache.get("content_length"),
        "content_type": _cache.get("content_type"),
        "final_url": _cache.get("final_url"),
        "status_code": _cache.get("status_code"),
    }
