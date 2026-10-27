"""
Bot de Telegram integrado al sitio Zoologic.

Lee directamente la planilla viva del sitio (DATA_DIR/zoologic_actual.xlsx),
la misma que se sube/edita desde la web. Cada consulta chequea si el archivo
cambió (fecha de modificación) y, si cambió, lo vuelve a leer. No hace falta
subir una planilla aparte ni usar /recargar.

Uso: escribirle al bot un sector, usuario o ubicación (ej: "ariel", "remoto",
"suipacha") y devuelve la línea completa de cada coincidencia.

Variables de entorno:
    TELEGRAM_BOT_TOKEN  token de @BotFather (obligatorio)
    ALLOWED_CHAT_IDS    chat_id autorizados, separados por coma (muy recomendado)
    DATA_DIR            mismo volumen que el sitio (default /app/data)
    SEARCH_FIELDS       campos donde busca, separados por coma
                        (default: sector,usuario,ubicacion)
    SHOW_PASSWORD       "1" para mostrar columnas de password (default: oculto)
"""

from __future__ import annotations

import html
import logging
import os
import sys
import threading
from datetime import datetime
from typing import Any

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

# Reutiliza la lógica del sitio (lectura de planilla, detección de encabezados)
import app as site

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
SHOW_PASSWORD = os.environ.get("SHOW_PASSWORD", "0").strip() == "1"
SEARCH_FIELDS = [
    f.strip().lower()
    for f in os.environ.get("SEARCH_FIELDS", "sector,usuario,ubicacion").split(",")
    if f.strip()
]
_allowed = os.environ.get("ALLOWED_CHAT_IDS", "").strip()
ALLOWED_CHAT_IDS = {int(x) for x in _allowed.split(",") if x.strip()} if _allowed else set()

MAX_MSG_LEN = 3800  # Telegram permite 4096
MAX_RESULTS = 15  # si hay más, pide afinar la búsqueda

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger("zoologic_bot")


# --------------------------------------------------------------------------
# Lectura de la planilla (con caché por fecha de modificación)
# --------------------------------------------------------------------------

_cache: dict[str, Any] = {"sig": None, "data": None}
_lock = threading.Lock()


def _signature() -> tuple[float, int] | None:
    try:
        st = site.CURRENT_FILE.stat()
    except FileNotFoundError:
        return None
    return (st.st_mtime, st.st_size)


def _field_headers(headers: list[dict[str, Any]], field: str) -> list[dict[str, Any]]:
    """Columnas que corresponden a un campo de búsqueda (sector, usuario, ubicacion...)."""
    found = []
    for h in headers:
        label = site.normalize_text(h["label"])
        if field == "usuario":
            ok = "usuario" in label and not label.startswith("fecha")
        else:
            ok = label.startswith(field)
        if ok:
            found.append(h)
    return found


def _load() -> dict[str, Any]:
    wb = site.open_workbook()
    try:
        sheet = wb[site.preferred_sheet_name(wb)]
        header_row = site.detect_header_row(sheet)
        headers = site.read_headers(sheet, header_row)
        rows = []
        for n in range(header_row + 1, sheet.max_row + 1):
            values = {h["key"]: site.serialize_value(sheet.cell(n, h["column"]).value) for h in headers}
            if any(v not in (None, "") for v in values.values()):
                rows.append({"row_id": n, "values": values})
        search_headers = []
        for field in SEARCH_FIELDS:
            search_headers += _field_headers(headers, field)
        for r in rows:
            r["haystack"] = " | ".join(
                site.normalize_text(r["values"].get(h["key"])) for h in search_headers
            )
        return {
            "sheet": sheet.title,
            "headers": headers,
            "rows": rows,
            "search_labels": [h["label"] for h in search_headers],
        }
    finally:
        wb.close()


def get_data() -> tuple[dict[str, Any], tuple[float, int]]:
    """Devuelve la planilla, releyéndola solo si el archivo cambió."""
    sig = _signature()
    if sig is None:
        raise FileNotFoundError("Todavía no hay una planilla cargada en el sitio.")
    with _lock:
        if _cache["sig"] != sig:
            try:
                data = _load()
            except Exception:
                # el sitio puede estar guardando justo en este momento: un reintento
                logger.warning("Lectura fallida, reintento", exc_info=True)
                import time

                time.sleep(1.5)
                sig = _signature() or sig
                data = _load()
            _cache["sig"], _cache["data"] = sig, data
            logger.info("Planilla (re)cargada: %d filas", len(data["rows"]))
        return _cache["data"], _cache["sig"]


# --------------------------------------------------------------------------
# Búsqueda y formato
# --------------------------------------------------------------------------

def buscar(texto: str) -> tuple[list[dict[str, Any]], dict[str, Any], tuple[float, int]]:
    data, sig = get_data()
    key = site.normalize_text(texto)
    if not key:
        return [], data, sig
    return [r for r in data["rows"] if key in r["haystack"]], data, sig


def _es_password(label: str) -> bool:
    t = site.normalize_text(label)
    return "password" in t or "contrasena" in t or "clave" in t


def _valor(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def formatear_fila(row: dict[str, Any], headers: list[dict[str, Any]]) -> str:
    lineas = []
    for h in headers:
        if not SHOW_PASSWORD and _es_password(h["label"]):
            continue
        val = _valor(row["values"].get(h["key"]))
        if val:
            label = h["label"]
            if len(label) > 40:  # p. ej. el encabezado larguísimo de "Puesto"
                label = label.split("(")[0].strip()
            lineas.append(f"<b>{html.escape(label)}:</b> {html.escape(val)}")
    return "\n".join(lineas)


def fecha_planilla(sig: tuple[float, int]) -> str:
    return datetime.fromtimestamp(sig[0]).strftime("%d/%m/%Y %H:%M")


def armar_mensajes(rows: list[dict[str, Any]], data: dict[str, Any], sig: tuple[float, int]) -> list[str]:
    pie = f"\n\n<i>Planilla actualizada: {fecha_planilla(sig)}</i>"
    if not rows:
        return ["No encontré ninguna coincidencia." + pie]
    extra = ""
    if len(rows) > MAX_RESULTS:
        extra = f"\n(Muestro los primeros {MAX_RESULTS}; afiná la búsqueda para ver el resto.)"
        rows = rows[:MAX_RESULTS]
    total = len(rows)
    bloques = [formatear_fila(r, data["headers"]) for r in rows]
    mensajes, actual = [], f"Encontré <b>{total}</b> resultado(s):{extra}"
    for b in bloques:
        b = b[:MAX_MSG_LEN]
        if len(actual) + len(b) + 2 > MAX_MSG_LEN:
            mensajes.append(actual)
            actual = b
        else:
            actual += "\n\n" + b
    mensajes.append(actual)
    if len(mensajes[-1]) + len(pie) > 4000:
        mensajes.append(pie.strip())
    else:
        mensajes[-1] += pie
    return mensajes


# --------------------------------------------------------------------------
# Handlers
# --------------------------------------------------------------------------

def _autorizado(update: Update) -> bool:
    return bool(ALLOWED_CHAT_IDS) and update.effective_chat.id in ALLOWED_CHAT_IDS


async def _rechazar(update: Update) -> None:
    await update.message.reply_text(
        f"No estás autorizado para usar este bot. Pasale este ID al administrador: {update.effective_chat.id}"
    )


async def cmd_id(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(f"Tu chat_id es: {update.effective_chat.id}")


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _autorizado(update):
        return await _rechazar(update)
    await update.message.reply_text(
        "Hola! Escribime un sector, usuario o ubicación (ej: ariel, remoto, suipacha) "
        "y te traigo la línea completa de la planilla.\n\n"
        "Siempre consulto la planilla actual del sitio, no hace falta recargar nada.\n\n"
        "/estado - fecha de la planilla y cantidad de filas\n"
        "/cambios - últimos cambios hechos desde el sitio\n"
        "/id - muestra tu chat_id"
    )


async def cmd_estado(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _autorizado(update):
        return await _rechazar(update)
    try:
        data, sig = get_data()
    except Exception as e:
        return await update.message.reply_text(f"No pude leer la planilla: {e}")
    await update.message.reply_text(
        f"Hoja: {data['sheet']}\nFilas: {len(data['rows'])}\n"
        f"Última modificación: {fecha_planilla(sig)}\n"
        f"Busco en: {', '.join(data['search_labels']) or '(sin columnas)'}"
    )


async def cmd_cambios(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _autorizado(update):
        return await _rechazar(update)
    entradas = site.read_history(8)
    if not entradas:
        return await update.message.reply_text("Todavía no hay cambios registrados desde el sitio.")
    partes = []
    for e in entradas:
        cab = f"<b>{html.escape(e['timestamp'].replace('T', ' '))}</b> · fila {e['row_id']}"
        det = [
            f"• {html.escape(str(c['field']))}: {html.escape(_valor(c['old']) or '(vacío)')} → "
            f"{html.escape(_valor(c['new']) or '(vacío)')}"
            for c in e["changes"]
        ]
        partes.append(cab + "\n" + "\n".join(det))
    texto = "\n\n".join(partes)[:MAX_MSG_LEN]
    await update.message.reply_text(texto, parse_mode=ParseMode.HTML)


async def manejar_texto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _autorizado(update):
        return await _rechazar(update)
    try:
        rows, data, sig = buscar(update.message.text or "")
    except Exception as e:
        logger.exception("Error buscando")
        return await update.message.reply_text(f"Ocurrió un error buscando: {e}")
    for msg in armar_mensajes(rows, data, sig):
        await update.message.reply_text(msg, parse_mode=ParseMode.HTML)


def main() -> None:
    if not TOKEN:
        sys.exit("Falta la variable de entorno TELEGRAM_BOT_TOKEN.")
    if not ALLOWED_CHAT_IDS:
        logger.warning("ALLOWED_CHAT_IDS vacío: el bot rechaza a todos (usá /id para obtener tu chat_id).")
    tg = Application.builder().token(TOKEN).build()
    tg.add_handler(CommandHandler("start", cmd_start))
    tg.add_handler(CommandHandler("id", cmd_id))
    tg.add_handler(CommandHandler("estado", cmd_estado))
    tg.add_handler(CommandHandler("cambios", cmd_cambios))
    tg.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, manejar_texto))
    logger.info("Bot iniciado. Esperando mensajes...")
    tg.run_polling()


if __name__ == "__main__":
    main()
