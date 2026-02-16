from flask import Flask, render_template, request, jsonify
import openpyxl
import os
import time

from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

app = Flask(__name__)

ONEDRIVE_SHARE_URL = os.environ.get("ONEDRIVE_SHARE_URL", "").strip()
XLSM_PATH = os.path.join(os.path.dirname(__file__), "..", "ebayマスタコピー.xlsm")
SHEET_NAME = "仕入・在庫管理表"
HEADER_ROW = 5
DATA_START_ROW = 6
DATA_END_ROW = 4847
MAX_COL = 28  # A〜AB

# キャッシュ: {rows_data: {行番号: [値リスト]}, headers: [...], mtime, source}
_cache = {"rows_data": None, "headers": None, "mtime": None, "source": None}


def _parse_workbook(wb):
    """workbook から全行データとヘッダーを抽出してメモリにキャッシュする形式で返す"""
    ws = wb[SHEET_NAME]

    # read_only モードでは iter_rows で一括読み込みが最速
    rows_data = {}
    for row in ws.iter_rows(min_row=HEADER_ROW, max_row=DATA_END_ROW, max_col=MAX_COL, values_only=False):
        r = row[0].row
        rows_data[r] = [cell.value for cell in row]

    # ヘッダー構築
    header_values = rows_data.get(HEADER_ROW, [None] * MAX_COL)
    headers = []
    for col in range(MAX_COL):
        col_letter = openpyxl.utils.get_column_letter(col + 1)
        val = header_values[col] if col < len(header_values) else None
        headers.append({"col": col + 1, "letter": col_letter, "name": val or f"({col_letter})"})

    wb.close()
    return rows_data, headers


def get_data():
    """行データとヘッダーを返す（キャッシュ付き）"""
    if ONEDRIVE_SHARE_URL:
        return _get_data_onedrive()
    return _get_data_local()


def _get_data_onedrive():
    """OneDrive 版: read_only + メモリキャッシュ"""
    from onedrive import fetch_xlsm, _is_cache_fresh

    if _cache["source"] == "onedrive" and _cache["rows_data"] is not None and _is_cache_fresh():
        return _cache["rows_data"], _cache["headers"]

    buf = fetch_xlsm(ONEDRIVE_SHARE_URL)
    wb = openpyxl.load_workbook(buf, data_only=True, keep_vba=False, read_only=True)
    rows_data, headers = _parse_workbook(wb)
    _cache.update(rows_data=rows_data, headers=headers, mtime=None, source="onedrive")
    return rows_data, headers


def _refresh_onedrive():
    """OneDrive キャッシュを強制クリアして再取得"""
    from onedrive import invalidate_cache
    invalidate_cache()
    _cache.update(rows_data=None, headers=None, mtime=None, source=None)


def _get_data_local():
    """ローカルファイル版: mtime ベースのキャッシュ"""
    mtime = os.path.getmtime(XLSM_PATH)
    if _cache["mtime"] != mtime or _cache["source"] != "local":
        wb = openpyxl.load_workbook(XLSM_PATH, data_only=True, keep_vba=False, read_only=True)
        rows_data, headers = _parse_workbook(wb)
        _cache.update(rows_data=rows_data, headers=headers, mtime=mtime, source="local")
    return _cache["rows_data"], _cache["headers"]


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """キャッシュをクリアして最新データを再取得する"""
    if ONEDRIVE_SHARE_URL:
        _refresh_onedrive()
        t0 = time.time()
        get_data()
        elapsed = time.time() - t0
        return jsonify({"status": "refreshed", "source": "onedrive", "elapsed_s": round(elapsed, 2)})
    return jsonify({"status": "skipped", "source": "local", "message": "ローカルモードではmtimeで自動更新されます"})


@app.route("/api/headers")
def api_headers():
    _, headers = get_data()
    return jsonify(headers)


@app.route("/api/cell")
def get_cell():
    try:
        col = int(request.args.get("col", 1))
        row = int(request.args.get("row", DATA_START_ROW))
    except (TypeError, ValueError):
        return jsonify({"error": "col と row は整数で指定してください"}), 400

    if col < 1 or col > MAX_COL:
        return jsonify({"error": f"col は 1〜{MAX_COL} の範囲で指定してください"}), 400
    if row < DATA_START_ROW or row > DATA_END_ROW:
        return jsonify({"error": f"row は {DATA_START_ROW}〜{DATA_END_ROW} の範囲で指定してください"}), 400

    rows_data, headers = get_data()
    row_values = rows_data.get(row)
    value = row_values[col - 1] if row_values and col - 1 < len(row_values) else None
    return jsonify({
        "col": col,
        "row": row,
        "header": headers[col - 1]["name"],
        "value": str(value) if value is not None else None,
    })


@app.route("/api/range")
def get_range():
    try:
        row_start = int(request.args.get("row_start", DATA_START_ROW))
        row_end = int(request.args.get("row_end", DATA_START_ROW + 19))
    except (TypeError, ValueError):
        return jsonify({"error": "row_start と row_end は整数で指定してください"}), 400

    row_start = max(row_start, DATA_START_ROW)
    row_end = min(row_end, DATA_END_ROW)

    if row_end - row_start > 100:
        row_end = row_start + 100

    rows_data, headers = get_data()

    rows = []
    for r in range(row_start, row_end + 1):
        row_data = {}
        row_values = rows_data.get(r)
        for col in range(MAX_COL):
            val = row_values[col] if row_values and col < len(row_values) else None
            row_data[headers[col]["letter"]] = str(val) if val is not None else None
        rows.append({"row": r, "data": row_data})

    return jsonify({
        "row_start": row_start,
        "row_end": row_end,
        "headers": headers,
        "rows": rows,
    })


if __name__ == "__main__":
    source = "OneDrive" if ONEDRIVE_SHARE_URL else f"Local ({XLSM_PATH})"
    print(f"Source: {source}")
    print(f"Sheet: {SHEET_NAME}")
    app.run(debug=True, host="0.0.0.0", port=5000)
