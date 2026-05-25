from __future__ import annotations

import os
import shutil
import unicodedata
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import load_workbook
from werkzeug.utils import secure_filename


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR / "data"))
UPLOAD_DIR = DATA_DIR / "uploads"
BACKUP_DIR = DATA_DIR / "backups"
CURRENT_FILE = DATA_DIR / "zoologic_actual.xlsx"
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)


def normalize_text(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(text.casefold().split())


def serialize_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    return value


def workbook_path() -> Path:
    if not CURRENT_FILE.exists():
        raise FileNotFoundError("Todavia no hay una planilla cargada.")
    return CURRENT_FILE


def open_workbook():
    return load_workbook(workbook_path())


def header_score(values: list[Any]) -> int:
    normalized = [normalize_text(value) for value in values if value not in (None, "")]
    text = " | ".join(normalized)
    score = len(normalized)
    for keyword in ("usuario", "puesto", "serie", "concepto", "canal", "estado"):
        if keyword in text:
            score += 8
    if "usuario" in text and "puesto" in text:
        score += 30
    return score


def detect_header_row(sheet) -> int:
    best_row = 1
    best_score = -1
    limit = min(sheet.max_row, 15)
    for row_number in range(1, limit + 1):
        values = [sheet.cell(row_number, col).value for col in range(1, sheet.max_column + 1)]
        score = header_score(values)
        if score > best_score:
            best_row = row_number
            best_score = score
    return best_row


def read_headers(sheet, header_row: int) -> list[dict[str, Any]]:
    headers: list[dict[str, Any]] = []
    seen: dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        raw_value = sheet.cell(header_row, column).value
        if raw_value in (None, ""):
            continue
        label = str(raw_value).strip()
        count = seen.get(label, 0) + 1
        seen[label] = count
        key = label if count == 1 else f"{label} ({count})"
        headers.append({"key": key, "label": label, "column": column})
    return headers


def searchable_headers(headers: list[dict[str, Any]]) -> list[dict[str, Any]]:
    matches = []
    for header in headers:
        text = normalize_text(header["label"])
        is_user = "usuario" in text and not text.startswith("fecha")
        is_position = "puesto" in text
        if is_user or is_position:
            matches.append(header)
    return matches


def preferred_sheet_name(workbook) -> str:
    for sheet in workbook.worksheets:
        header_row = detect_header_row(sheet)
        headers = read_headers(sheet, header_row)
        labels = " | ".join(normalize_text(header["label"]) for header in headers)
        if "usuario" in labels and "puesto" in labels:
            return sheet.title
    return workbook.active.title


def sheet_names(workbook) -> list[str]:
    return [sheet.title for sheet in workbook.worksheets]


def rows_from_sheet(sheet_name: str | None = None, search: str = "") -> dict[str, Any]:
    workbook = open_workbook()
    try:
        selected_sheet = sheet_name if sheet_name in workbook.sheetnames else preferred_sheet_name(workbook)
        sheet = workbook[selected_sheet]
        header_row = detect_header_row(sheet)
        headers = read_headers(sheet, header_row)
        search_headers = searchable_headers(headers)
        search_key = normalize_text(search)
        rows = []

        for row_number in range(header_row + 1, sheet.max_row + 1):
            values = {}
            has_value = False
            for header in headers:
                value = serialize_value(sheet.cell(row_number, header["column"]).value)
                values[header["key"]] = value
                if value not in (None, ""):
                    has_value = True

            if not has_value:
                continue

            if search_key:
                haystack_headers = search_headers or headers
                haystack = " | ".join(
                    normalize_text(sheet.cell(row_number, header["column"]).value)
                    for header in haystack_headers
                )
                if search_key not in haystack:
                    continue

            rows.append({"row_id": row_number, "values": values})

        return {
            "filename": CURRENT_FILE.name,
            "sheet": sheet.title,
            "sheets": sheet_names(workbook),
            "header_row": header_row,
            "headers": headers,
            "search_columns": [header["label"] for header in search_headers],
            "rows": rows,
            "total_rows": max(sheet.max_row - header_row, 0),
            "filtered_rows": len(rows),
        }
    finally:
        workbook.close()


def make_backup() -> None:
    if CURRENT_FILE.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(CURRENT_FILE, BACKUP_DIR / f"zoologic_backup_{stamp}.xlsx")


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/status")
def status():
    if not CURRENT_FILE.exists():
        return jsonify({"loaded": False})
    data = rows_from_sheet()
    return jsonify(
        {
            "loaded": True,
            "filename": data["filename"],
            "sheet": data["sheet"],
            "sheets": data["sheets"],
            "headers": data["headers"],
            "search_columns": data["search_columns"],
            "total_rows": data["total_rows"],
        }
    )


@app.post("/api/upload")
def upload_file():
    ensure_dirs()
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "Subi un archivo Excel para continuar."}), 400

    suffix = Path(file.filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Formato no soportado. Usa .xlsx o .xlsm."}), 400

    filename = secure_filename(file.filename)
    uploaded_path = UPLOAD_DIR / filename
    file.save(uploaded_path)

    try:
        workbook = load_workbook(uploaded_path)
        workbook.close()
    except Exception as exc:
        uploaded_path.unlink(missing_ok=True)
        return jsonify({"error": f"No pude abrir el Excel: {exc}"}), 400

    make_backup()
    shutil.copy2(uploaded_path, CURRENT_FILE)
    return jsonify(rows_from_sheet())


@app.get("/api/rows")
def rows():
    try:
        return jsonify(rows_from_sheet(request.args.get("sheet"), request.args.get("q", "")))
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404


@app.patch("/api/rows/<int:row_id>")
def update_row(row_id: int):
    payload = request.get_json(silent=True) or {}
    values = payload.get("values", {})
    sheet_name = payload.get("sheet")
    if not isinstance(values, dict):
        return jsonify({"error": "Los datos enviados no son validos."}), 400

    try:
        workbook = open_workbook()
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404

    try:
        selected_sheet = sheet_name if sheet_name in workbook.sheetnames else preferred_sheet_name(workbook)
        sheet = workbook[selected_sheet]
        header_row = detect_header_row(sheet)
        headers = read_headers(sheet, header_row)
        key_to_column = {header["key"]: header["column"] for header in headers}

        if row_id <= header_row or row_id > sheet.max_row:
            return jsonify({"error": "La fila indicada no existe."}), 404

        unknown = [key for key in values if key not in key_to_column]
        if unknown:
            return jsonify({"error": f"Columnas desconocidas: {', '.join(unknown)}"}), 400

        make_backup()
        for key, value in values.items():
            sheet.cell(row=row_id, column=key_to_column[key]).value = value if value != "" else None

        workbook.save(CURRENT_FILE)
        return jsonify({"ok": True})
    finally:
        workbook.close()


@app.get("/api/download")
def download():
    try:
        path = workbook_path()
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    return send_file(path, as_attachment=True, download_name="zoologic_actualizado.xlsx")


if __name__ == "__main__":
    ensure_dirs()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=False)
