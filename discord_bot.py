"""
Bot de Discord para consultar la planilla Zoologic.

Usa la misma lectura de planilla que el bot de Telegram (bot.py): lee la
planilla viva del sitio y la relee sola cuando cambia.

Uso: escribirle al bot (por mensaje directo o en un canal del servidor) un
sector, usuario o ubicación y devuelve la línea completa de cada coincidencia.

Variables de entorno:
    DISCORD_BOT_TOKEN          token del bot (obligatorio)
    DISCORD_ALLOWED_USER_IDS   IDs de usuario autorizados, separados por coma
    DISCORD_CHANNEL_IDS        opcional: solo responde en estos canales (y en DM)
    (más las de bot.py: DATA_DIR, SEARCH_FIELDS, SHOW_PASSWORD)

Comandos: !id  !estado  !cambios
"""

from __future__ import annotations

import logging
import os
import sys
from typing import Any

import discord

import app as site
import bot as core  # lectura de planilla y búsqueda compartidas

TOKEN = os.environ.get("DISCORD_BOT_TOKEN", "").strip()


def _ids(name: str) -> set[int]:
    raw = os.environ.get(name, "").strip()
    return {int(x) for x in raw.split(",") if x.strip()} if raw else set()


ALLOWED_USER_IDS = _ids("DISCORD_ALLOWED_USER_IDS")
CHANNEL_IDS = _ids("DISCORD_CHANNEL_IDS")
MAX_LEN = 1900  # Discord permite 2000 por mensaje

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger("zoologic_discord")


def formatear_fila(row: dict[str, Any], headers: list[dict[str, Any]]) -> str:
    lineas = []
    for h in headers:
        if not core.SHOW_PASSWORD and core._es_password(h["label"]):
            continue
        val = core._valor(row["values"].get(h["key"]))
        if val:
            label = h["label"]
            if len(label) > 40:
                label = label.split("(")[0].strip()
            lineas.append(f"**{discord.utils.escape_markdown(label)}:** {discord.utils.escape_markdown(val)}")
    return "\n".join(lineas)


def armar_mensajes(rows, data, sig) -> list[str]:
    pie = f"\n\n*Planilla actualizada: {core.fecha_planilla(sig)}*"
    if not rows:
        return ["No encontré ninguna coincidencia." + pie]
    extra = ""
    if len(rows) > core.MAX_RESULTS:
        extra = f"\n(Muestro los primeros {core.MAX_RESULTS}; afiná la búsqueda para ver el resto.)"
        rows = rows[: core.MAX_RESULTS]
    mensajes, actual = [], f"Encontré **{len(rows)}** resultado(s):{extra}"
    for r in rows:
        b = formatear_fila(r, data["headers"])[:MAX_LEN]
        if len(actual) + len(b) + 2 > MAX_LEN:
            mensajes.append(actual)
            actual = b
        else:
            actual += "\n\n" + b
    mensajes.append(actual)
    if len(mensajes[-1]) + len(pie) > 2000:
        mensajes.append(pie.strip())
    else:
        mensajes[-1] += pie
    return mensajes


intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)
NO_MENTIONS = discord.AllowedMentions.none()  # que un dato con @everyone no avise a nadie


async def responder(message: discord.Message, texto: str) -> None:
    await message.channel.send(texto, allowed_mentions=NO_MENTIONS)


@client.event
async def on_ready():
    logger.info("Bot de Discord conectado como %s", client.user)


@client.event
async def on_message(message: discord.Message):
    if message.author.bot or message.author == client.user:
        return
    es_dm = message.guild is None
    if not es_dm and CHANNEL_IDS and message.channel.id not in CHANNEL_IDS:
        return
    texto = (message.content or "").strip()
    if not texto:
        return

    if texto.lower() == "!id":
        return await responder(message, f"Tu ID de usuario es: {message.author.id}")

    if message.author.id not in ALLOWED_USER_IDS:
        if es_dm:  # en canales del servidor se ignora en silencio
            await responder(message, f"No estás autorizado. Pasale este ID al administrador: {message.author.id}")
        return

    cmd = texto.lower()
    try:
        if cmd == "!estado":
            data, sig = core.get_data()
            return await responder(
                message,
                f"Hoja: {data['sheet']}\nFilas: {len(data['rows'])}\n"
                f"Última modificación: {core.fecha_planilla(sig)}\n"
                f"Busco en: {', '.join(data['search_labels']) or '(sin columnas)'}",
            )
        if cmd == "!cambios":
            entradas = site.read_history(8)
            if not entradas:
                return await responder(message, "Todavía no hay cambios registrados desde el sitio.")
            partes = []
            for e in entradas:
                det = [
                    f"• {discord.utils.escape_markdown(str(c['field']))}: "
                    f"{discord.utils.escape_markdown(core._valor(c['old']) or '(vacío)')} → "
                    f"{discord.utils.escape_markdown(core._valor(c['new']) or '(vacío)')}"
                    for c in e["changes"]
                ]
                partes.append(f"**{e['timestamp'].replace('T', ' ')}** · fila {e['row_id']}\n" + "\n".join(det))
            return await responder(message, "\n\n".join(partes)[:MAX_LEN])
        if cmd.startswith("!"):
            return
        async with message.channel.typing():
            rows, data, sig = core.buscar(texto)
        for msg in armar_mensajes(rows, data, sig):
            await responder(message, msg)
    except Exception as e:
        logger.exception("Error procesando mensaje")
        await responder(message, f"Ocurrió un error: {e}")


def main() -> None:
    if not TOKEN:
        sys.exit("Falta la variable de entorno DISCORD_BOT_TOKEN.")
    if not ALLOWED_USER_IDS:
        logger.warning("DISCORD_ALLOWED_USER_IDS vacío: el bot rechaza a todos (escribile !id para ver tu ID).")
    client.run(TOKEN, log_handler=None)


if __name__ == "__main__":
    main()
