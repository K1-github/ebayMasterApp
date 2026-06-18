"""OneDrive 共有リンク経由で xlsm ファイルをダウンロードする（読み取り専用）

1drv.ms 短縮URLを解決し、onedrive.live.com の直接ダウンロードURLを構築する。
"""

import base64
import io
import re
import time
from urllib.parse import urlencode, urljoin, urlparse, parse_qs, unquote

import requests


# --- キャッシュ ---
_cache = {"data": None, "fetched_at": 0.0, "filename": None, "content_length": None}
CACHE_TTL = 300  # 5 分


def _is_cache_fresh() -> bool:
    if _cache["data"] is None:
        return False
    return (time.time() - _cache["fetched_at"]) < CACHE_TTL


def invalidate_cache():
    """キャッシュを強制クリアする（次回 fetch 時に再ダウンロード）"""
    _cache.update(data=None, fetched_at=0.0)


def _make_session():
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        "Cache-Control": "no-cache",
    })
    return session


def _resolve_short_url(share_url: str, session: requests.Session) -> str:
    """1drv.ms 短縮URLをリダイレクト追跡して最終URLを返す"""
    resp = session.head(share_url, allow_redirects=True, timeout=30)
    return resp.url


def _build_direct_download_url(resolved_url: str) -> str:
    """onedrive.live.com の共有URLから直接ダウンロードURLを構築する"""
    parsed = urlparse(resolved_url)
    params = parse_qs(parsed.query)

    # resid と authkey を取得
    resid = params.get("resid", [None])[0]
    authkey = params.get("authkey", [None])[0]

    if resid and authkey:
        return f"https://onedrive.live.com/download?resid={resid}&authkey={authkey}&em=2"

    # cid パラメータがある場合
    cid = params.get("cid", [None])[0]
    if resid and cid:
        return f"https://onedrive.live.com/download?resid={resid}&cid={cid}&em=2"

    return None


def _try_graph_api_download(share_url: str, session: requests.Session) -> requests.Response:
    """Microsoft Graph API の sharing token 経由でダウンロードを試みる"""
    # sharing token = base64url("u!" + share_url)
    token = base64.urlsafe_b64encode(("u!" + share_url).encode()).decode().rstrip("=")
    api_url = f"https://api.onedrive.com/v1.0/shares/{token}/root/content"
    resp = session.get(api_url, allow_redirects=True, timeout=60)
    return resp


def fetch_xlsm(share_url: str) -> io.BytesIO:
    """共有リンクから xlsm をダウンロードし BytesIO で返す（キャッシュ付き）"""
    if _is_cache_fresh():
        _cache["data"].seek(0)
        return _cache["data"]

    session = _make_session()
    last_error = None

    # --- 方法1: 短縮URL解決 → 直接ダウンロードURL構築 ---
    try:
        resolved = _resolve_short_url(share_url, session)
        direct_url = _build_direct_download_url(resolved)
        if direct_url:
            resp = session.get(direct_url, allow_redirects=True, timeout=60)
            resp.raise_for_status()
            content_type = resp.headers.get("Content-Type", "")
            if "html" not in content_type and len(resp.content) > 1000:
                return _store_cache(resp)
            last_error = f"方法1失敗: content_type={content_type}, size={len(resp.content)}"
    except Exception as e:
        last_error = f"方法1例外: {e}"

    # --- 方法2: Graph API sharing token ---
    try:
        resp = _try_graph_api_download(share_url, session)
        resp.raise_for_status()
        content_type = resp.headers.get("Content-Type", "")
        if "html" not in content_type and len(resp.content) > 1000:
            return _store_cache(resp)
        last_error = f"方法2失敗: content_type={content_type}, size={len(resp.content)}"
    except Exception as e:
        last_error = f"方法2例外: {e}"

    # --- 方法3: download=1 を直接付与（元の方式） ---
    try:
        sep = "&" if "?" in share_url else "?"
        download_url = share_url + sep + "download=1" + f"&_t={int(time.time())}"
        resp = session.get(download_url, allow_redirects=True, timeout=60)
        resp.raise_for_status()
        content_type = resp.headers.get("Content-Type", "")
        if "html" not in content_type and len(resp.content) > 1000:
            return _store_cache(resp)
        last_error = f"方法3失敗: content_type={content_type}, size={len(resp.content)}"
    except Exception as e:
        last_error = f"方法3例外: {e}"

    raise ValueError(f"OneDriveからのダウンロードに失敗しました。最後のエラー: {last_error}")


def _store_cache(resp: requests.Response) -> io.BytesIO:
    """レスポンスをキャッシュに保存して BytesIO を返す"""
    content_type = resp.headers.get("Content-Type", "")
    filename = None
    cd = resp.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*=UTF-8''(.+?)(?:;|$)", cd)
    if m:
        filename = unquote(m.group(1))
    else:
        m = re.search(r'filename="?([^";]+)"?', cd)
        if m:
            filename = m.group(1)

    buf = io.BytesIO(resp.content)
    _cache.update(
        data=buf, fetched_at=time.time(), filename=filename,
        content_length=len(resp.content), content_type=content_type,
        final_url=resp.url, status_code=resp.status_code,
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
