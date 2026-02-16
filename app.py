from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from datetime import datetime, timezone, timedelta
from functools import wraps
import glob
import openpyxl
import os
import time

from dotenv import load_dotenv

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

# シート設定: {シート名: {header_row, data_start, search_col (0=A,1=B), max_col}}
SHEETS = {
    "仕入・在庫管理表": {"header_row": 5, "data_start": 6, "search_col": 1, "search_label": "出品管理ID（B列）", "max_col": 28},
    "販売管理表": {"header_row": 5, "data_start": 6, "search_col": 0, "search_label": "レコード番号（A列）", "max_col": 28},
    "無在庫管理表（中古）": {"header_row": 5, "data_start": 6, "search_col": 1, "search_label": "出品管理ID（B列）", "max_col": 28},
}

# シートごとのキャッシュ
_cache = {}
_wb_cache = {"mtime": None, "source": None}


def _parse_sheet(wb, sheet_name):
    """指定シートのデータとヘッダーを読み取る"""
    cfg = SHEETS[sheet_name]
    ws = wb[sheet_name]
    header_row = cfg["header_row"]
    max_col = cfg["max_col"]

    rows_data = {}
    max_data_row = header_row
    for row in ws.iter_rows(min_row=header_row, max_col=max_col, values_only=False):
        r = row[0].row
        values = [cell.value for cell in row]
        if any(v is not None for v in values):
            rows_data[r] = values
            max_data_row = r

    header_values = rows_data.get(header_row, [None] * max_col)
    headers = []
    for col in range(max_col):
        col_letter = openpyxl.utils.get_column_letter(col + 1)
        val = header_values[col] if col < len(header_values) else None
        headers.append({"col": col + 1, "letter": col_letter, "name": val or f"({col_letter})"})

    return {"rows_data": rows_data, "headers": headers, "max_row": max_data_row}


def _load_all_sheets(wb):
    """全シートを読み込んでキャッシュに格納"""
    for sheet_name in SHEETS:
        if sheet_name in wb.sheetnames:
            _cache[sheet_name] = _parse_sheet(wb, sheet_name)
    wb.close()


def get_sheet_data(sheet_name):
    """指定シートのデータを返す（キャッシュ付き）"""
    if ONEDRIVE_SHARE_URL:
        return _get_data_onedrive(sheet_name)
    return _get_data_local(sheet_name)


def _get_data_onedrive(sheet_name):
    from onedrive import fetch_xlsm, _is_cache_fresh

    if _wb_cache["source"] == "onedrive" and sheet_name in _cache and _is_cache_fresh():
        c = _cache[sheet_name]
        return c["rows_data"], c["headers"], c["max_row"]

    buf = fetch_xlsm(ONEDRIVE_SHARE_URL)
    wb = openpyxl.load_workbook(buf, data_only=True, keep_vba=False, read_only=True)
    _load_all_sheets(wb)
    _wb_cache.update(mtime=None, source="onedrive")
    c = _cache[sheet_name]
    return c["rows_data"], c["headers"], c["max_row"]


def _refresh_onedrive():
    from onedrive import invalidate_cache
    invalidate_cache()
    _cache.clear()
    _wb_cache.update(mtime=None, source=None)


def _get_data_local(sheet_name):
    global XLSM_PATH
    if not XLSM_PATH:
        XLSM_PATH = _find_xlsm()
    if not XLSM_PATH:
        raise FileNotFoundError("ebayマスタ*.xlsm が見つかりません")
    mtime = os.path.getmtime(XLSM_PATH)
    if _wb_cache["mtime"] != mtime or _wb_cache["source"] != "local":
        wb = openpyxl.load_workbook(XLSM_PATH, data_only=True, keep_vba=False, read_only=True)
        _load_all_sheets(wb)
        _wb_cache.update(mtime=mtime, source="local")
    c = _cache[sheet_name]
    return c["rows_data"], c["headers"], c["max_row"]


def _to_str(val):
    """セル値を検索用文字列に変換"""
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
        _refresh_onedrive()
        t0 = time.time()
        # 全シート読み込み
        get_sheet_data(list(SHEETS.keys())[0])
        elapsed = time.time() - t0
        return jsonify({"status": "refreshed", "source": "onedrive", "elapsed_s": round(elapsed, 2)})
    return jsonify({"status": "skipped", "source": "local", "message": "ローカルモードではmtimeで自動更新されます"})


@app.route("/api/fileinfo")
@login_required
def api_fileinfo():
    JST = timezone(timedelta(hours=9))
    if ONEDRIVE_SHARE_URL:
        from onedrive import get_file_info
        info = get_file_info()
        fetched_at = None
        if info["fetched_at"]:
            fetched_at = datetime.fromtimestamp(info["fetched_at"], tz=JST).strftime("%Y-%m-%d %H:%M:%S JST")
        return jsonify({
            "source": "onedrive",
            "filename": info["filename"],
            "fetched_at": fetched_at,
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
    """指定シートの指定列で検索"""
    sheet = request.args.get("sheet", "").strip()
    query = request.args.get("q", "").strip()

    if sheet not in SHEETS:
        return jsonify({"error": "無効なシート名です"}), 400
    if not query:
        return jsonify({"error": "検索IDを入力してください"}), 400

    cfg = SHEETS[sheet]
    search_col = cfg["search_col"]
    data_start = cfg["data_start"]
    max_col = cfg["max_col"]

    rows_data, headers, max_row = get_sheet_data(sheet)

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
                row_data[headers[col]["letter"]] = str(v) if v is not None else None
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
    """利用可能なシート一覧を返す"""
    result = []
    for name, cfg in SHEETS.items():
        result.append({"name": name, "search_label": cfg["search_label"]})
    return jsonify(result)


if __name__ == "__main__":
    source = "OneDrive" if ONEDRIVE_SHARE_URL else f"Local ({XLSM_PATH})"
    print(f"Source: {source}")
    app.run(debug=True, host="0.0.0.0", port=5000)
