"""OneDrive 共有リンク経由で xlsm ファイルをダウンロードする（読み取り専用）

共有リンクに &download=1 を付与し、requests.Session でリダイレクトを
追跡することで認証なしで直接ダウンロードする。
"""

import io
import time

import requests


# --- キャッシュ ---
_cache = {"data": None, "fetched_at": 0.0}
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

    buf = io.BytesIO(resp.content)
    _cache.update(data=buf, fetched_at=time.time())
    buf.seek(0)
    return buf
