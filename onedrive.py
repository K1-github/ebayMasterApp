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

    session = requests.Session()
    download_url = _build_download_url(share_url)
    download_url += f"&_t={int(time.time())}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/octet-stream,*/*",
        "Cache-Control": "no-cache, no-store",
        "Pragma": "no-cache",
    }
    resp = session.get(download_url, headers=headers, timeout=60)
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "")
    if "html" in content_type:
        # HTML が返された場合: URLをリダイレクトで解決してから再試行
        resolved = session.get(share_url, headers=headers, timeout=30, allow_redirects=True)
        resolved_url = resolved.url
        # onedrive.live.com の直接ダウンロードURLに変換
        if "onedrive.live.com" in resolved_url or "sharepoint.com" in resolved_url:
            from urllib.parse import urlparse, parse_qs, urlencode
            parsed = urlparse(resolved_url)
            params = parse_qs(parsed.query)
            resid = params.get("resid", [None])[0]
            authkey = params.get("authkey", [None])[0]
            cid = params.get("cid", [None])[0]
            if resid and authkey:
                direct = f"https://onedrive.live.com/download?resid={resid}&authkey={authkey}&em=2"
                resp2 = session.get(direct, headers=headers, timeout=60)
                resp2.raise_for_status()
                content_type2 = resp2.headers.get("Content-Type", "")
                if "html" not in content_type2:
                    resp = resp2
                    content_type = content_type2

        if "html" in resp.headers.get("Content-Type", ""):
            raise ValueError(
                f"OneDriveがHTMLを返しました（共有URLが期限切れか無効の可能性）。"
                f"Content-Type: {resp.headers.get('Content-Type')}, "
                f"URL: {resp.url[:100]}"
            )

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
