from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from datetime import datetime, timezone, timedelta
from functools import wraps
import glob
import os
import time

from dotenv import load_dotenv

# Vercel上では環境変数はVercelダッシュボードで管理するため.envを読まない
if not os.environ.get("VERCEL"):
    load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-me")

APP_PASSWORD = os.environ.get("APP_PASSWORD", "")


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if APP_PASSWORD and not session.get("authenticated"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

ONEDRIVE_SHARE_URL = os.environ.get("ONEDRIVE_SHARE_URL", "").strip()

def _find_xlsm():
    pattern = os.path.join(os.path.dirname(__file__), "..", "ebayマスタ*.xlsm")
    matches = glob.glob(pattern)
    return matches[0] if matches else None

XLSM_PATH = _find_xlsm()

# シート設定: {キー: {sheet_tab, header_row, data_start, search_col (0=A,1=B,...), max_col}}
SHEETS = {
    "仕入・在庫管理表": {"header_row": 5, "data_start": 6, "search_col": 1, "search_label": "出品管理ID（B列）", "max_col": 28},
    "販売管理表": {"header_row": 5, "data_start": 6, "search_col": 0, "search_label": "レコード番号（A列）", "max_col": 28},
    "販売管理表_出品管理ID": {"sheet_tab": "販売管理表", "header_row": 5, "data_start": 6, "search_col": 1, "search_label": "出品管理ID（B列）", "max_col": 28},
    "無在庫管理表(中古)": {"header_row": 5, "data_start": 6, "search_col": 1, "search_label": "出品管理ID（B列）", "max_col": 28},
    "出品管理表": {"header_row": 6, "data_start": 7, "search_col": 2, "search_label": "出品管理ID（C列）", "max_col": 28},
    "DDP": {"header_row": 5, "data_start": 6, "search_col": 2, "search_label": "出品管理ID（C列）", "max_col": 28},
}

# シートごとのパース済みキャッシュ
_cache = {}
# ファイルレベルのキャッシュ状態
_wb_cache = {"mtime": None, "source": None, "file_ready": False}


def _col_letter(col_idx):
    """0始まりの列インデックスをExcel列名に変換"""
    n = col_idx + 1
    result = ""
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _parse_sheet_from_wb(wb, sheet_name):
    """calamine ワークブックから指定シートをパース"""
    cfg = SHEETS[sheet_name]
    tab = cfg.get("sheet_tab", sheet_name)
    header_row = cfg["header_row"]
    max_col = cfg["max_col"]
    data_start = cfg["data_start"]

    sheet = wb.get_sheet_by_name(tab)
    all_rows = sheet.to_python(skip_empty_area=False)

    rows_data = {}
    max_data_row = header_row

    for i, row_values in enumerate(all_rows):
        r = i + 1
        if r < header_row:
            continue
        values = list(row_values[:max_col]) + [None] * max(0, max_col - len(row_values))
        if any(v is not None for v in values):
            rows_data[r] = values
            if r >= data_start:
                max_data_row = r

    header_values = rows_data.get(header_row, [None] * max_col)
    headers = []
    for col in range(max_col):
        letter = _col_letter(col)
        val = header_values[col] if col < len(header_values) else None
        name = str(val) if val is not None else f"({letter})"
        headers.append({"col": col + 1, "letter": letter, "name": name})

    return {"rows_data": rows_data, "headers": headers, "max_row": max_data_row}


def _open_wb(buf_or_path):
    """calamine でワークブックを開く"""
    import tempfile
    from python_calamine import CalamineWorkbook
    if isinstance(buf_or_path, str):
        return CalamineWorkbook.from_path(buf_or_path)
    buf_or_path.seek(0)
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        tmp.write(buf_or_path.read())
        path = tmp.name
    return CalamineWorkbook.from_path(path)


def _load_all_sheets(wb):
    """全シートをパースしてキャッシュに格納"""
    available = set(wb.sheet_names)
    for sheet_name, cfg in SHEETS.items():
        tab = cfg.get("sheet_tab", sheet_name)
        if tab in available:
            _cache[sheet_name] = _parse_sheet_from_wb(wb, sheet_name)


def get_sheet_data(sheet_name):
    if ONEDRIVE_SHARE_URL:
        return _get_data_onedrive(sheet_name)
    return _get_data_local(sheet_name)


def _get_data_onedrive(sheet_name):
    from onedrive import fetch_xlsm, _is_cache_fresh

    # ファイルが新鮮かつシートがパース済みならキャッシュを返す
    if _wb_cache["source"] == "onedrive" and _wb_cache["file_ready"] and sheet_name in _cache and _is_cache_fresh():
        c = _cache[sheet_name]
        return c["rows_data"], c["headers"], c["max_row"]

    # ファイルが新鮮だがシートが未パース（遅延読み込み）
    if _wb_cache["source"] == "onedrive" and _wb_cache["file_ready"] and _is_cache_fresh() and sheet_name not in _cache:
        buf = fetch_xlsm(ONEDRIVE_SHARE_URL)
        wb = _open_wb(buf)
        _cache[sheet_name] = _parse_sheet_from_wb(wb, sheet_name)
        c = _cache[sheet_name]
        return c["rows_data"], c["headers"], c["max_row"]

    # ファイルを新規ダウンロード → 全シートパース
    buf = fetch_xlsm(ONEDRIVE_SHARE_URL)
    wb = _open_wb(buf)
    _load_all_sheets(wb)
    _wb_cache.update(mtime=None, source="onedrive", file_ready=True)
    c = _cache[sheet_name]
    return c["rows_data"], c["headers"], c["max_row"]


def _refresh_onedrive():
    """OneDriveキャッシュをクリアしてファイルだけ再ダウンロード（シートは遅延パース）"""
    from onedrive import invalidate_cache, fetch_xlsm
    invalidate_cache()
    _cache.clear()
    _wb_cache.update(mtime=None, source=None, file_ready=False)
    # ファイルをダウンロードしてキャッシュに乗せる（パースはしない）
    fetch_xlsm(ONEDRIVE_SHARE_URL)
    _wb_cache.update(source="onedrive", file_ready=True)


def _get_data_local(sheet_name):
    global XLSM_PATH
    if not XLSM_PATH:
        XLSM_PATH = _find_xlsm()
    if not XLSM_PATH:
        raise FileNotFoundError("ebayマスタ*.xlsm が見つかりません")
    mtime = os.path.getmtime(XLSM_PATH)
    if _wb_cache["mtime"] != mtime or _wb_cache["source"] != "local":
        wb = _open_wb(XLSM_PATH)
        _load_all_sheets(wb)
        _wb_cache.update(mtime=mtime, source="local", file_ready=True)
    c = _cache[sheet_name]
    return c["rows_data"], c["headers"], c["max_row"]


def _to_str(val):
    if val is None:
        return None
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val)


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("index"))
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["authenticated"] = True
            return redirect(url_for("index"))
        return render_template("login.html", error="パスワードが違います")
    return render_template("login.html")


@app.route("/")
@login_required
def index():
    return render_template("index.html")


@app.route("/api/refresh", methods=["POST"])
@login_required
def api_refresh():
    if ONEDRIVE_SHARE_URL:
        try:
            t0 = time.time()
            _refresh_onedrive()
            elapsed = time.time() - t0
            return jsonify({
                "status": "refreshed", "source": "onedrive",
                "elapsed_s": round(elapsed, 2),
            })
        except Exception as e:
            return jsonify({"status": "error", "error": str(e)}), 500
    return jsonify({"status": "skipped", "source": "local", "message": "ローカルモードではmtimeで自動更新されます"})


@app.route("/api/fileinfo")
@login_required
def api_fileinfo():
    JST = timezone(timedelta(hours=9))
    if ONEDRIVE_SHARE_URL:
        from onedrive import get_file_info, fetch_xlsm, _is_cache_fresh
        if not _is_cache_fresh():
            try:
                fetch_xlsm(ONEDRIVE_SHARE_URL)
            except Exception:
                pass
        info = get_file_info()
        fetched_at = None
        if info["fetched_at"]:
            fetched_at = datetime.fromtimestamp(info["fetched_at"], tz=JST).strftime("%Y-%m-%d %H:%M:%S JST")
        last_modified = None
        if info.get("last_modified"):
            last_modified = datetime.fromtimestamp(info["last_modified"], tz=JST).strftime("%Y-%m-%d %H:%M:%S JST")
        return jsonify({
            "source": "onedrive",
            "filename": info["filename"],
            "fetched_at": fetched_at,
            "last_modified": last_modified,
            "content_length": info["content_length"],
            "content_type": info.get("content_type"),
            "status_code": info.get("status_code"),
        })
    else:
        fname = os.path.basename(XLSM_PATH) if XLSM_PATH else None
        mtime = None
        if XLSM_PATH and os.path.exists(XLSM_PATH):
            mtime = datetime.fromtimestamp(os.path.getmtime(XLSM_PATH), tz=JST).strftime("%Y-%m-%d %H:%M:%S JST")
        return jsonify({"source": "local", "filename": fname, "last_modified": mtime})


@app.route("/api/search")
@login_required
def api_search():
    sheet = request.args.get("sheet", "").strip()
    query = request.args.get("q", "").strip()

    if sheet not in SHEETS:
        return jsonify({"error": f"無効なシート名です: {sheet}"}), 400
    if not query:
        return jsonify({"error": "検索IDを入力してください"}), 400

    try:
        cfg = SHEETS[sheet]
        search_col = cfg["search_col"]
        data_start = cfg["data_start"]
        max_col = cfg["max_col"]
        rows_data, headers, max_row = get_sheet_data(sheet)
    except Exception as e:
        return jsonify({"error": f"シート読み込みエラー: {sheet} - {str(e)}"}), 500

    matched = []
    for r in range(data_start, max_row + 1):
        row_values = rows_data.get(r)
        if not row_values:
            continue
        val = row_values[search_col] if search_col < len(row_values) else None
        val_str = _to_str(val)
        if val_str is not None and query in val_str.strip():
            row_data = {}
            for col in range(max_col):
                v = row_values[col] if col < len(row_values) else None
                row_data[headers[col]["letter"]] = _to_str(v)
            matched.append({"row": r, "data": row_data})

    return jsonify({
        "query": query,
        "sheet": sheet,
        "count": len(matched),
        "headers": headers,
        "rows": matched,
    })


@app.route("/api/sheets")
@login_required
def api_sheets():
    result = []
    for name, cfg in SHEETS.items():
        result.append({"name": name, "display_name": cfg.get("sheet_tab", name), "search_label": cfg["search_label"]})
    return jsonify(result)


@app.route("/api/debug-env")
def api_debug_env():
    return jsonify({
        "ONEDRIVE_SHARE_URL_set": bool(os.environ.get("ONEDRIVE_SHARE_URL")),
        "ONEDRIVE_SHARE_URL_len": len(os.environ.get("ONEDRIVE_SHARE_URL", "")),
        "dotenv_file_exists": os.path.exists(os.path.join(os.path.dirname(__file__), ".env")),
    })


if __name__ == "__main__":
    source = "OneDrive" if ONEDRIVE_SHARE_URL else f"Local ({XLSM_PATH})"
    print(f"Source: {source}")
    app.run(debug=True, host="0.0.0.0", port=5000)
