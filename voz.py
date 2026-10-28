"""
Búsqueda por voz: manda el audio al servicio Whisper (contenedor aparte) y
devuelve el texto. Sin dependencias extra (usa urllib).

Variables de entorno:
    WHISPER_URL        ej: http://192.168.1.50:9000  (vacío = voz desactivada)
    WHISPER_TOKEN      misma clave que en el servicio Whisper
    VOZ_MAX_SECONDS    duración máxima aceptada (default 60)
"""

from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request

WHISPER_URL = os.environ.get("WHISPER_URL", "").strip().rstrip("/")
WHISPER_TOKEN = os.environ.get("WHISPER_TOKEN", "").strip()
MAX_SECONDS = int(os.environ.get("VOZ_MAX_SECONDS", "60"))
MAX_BYTES = 10 * 1024 * 1024
TIMEOUT = 180  # la primera vez el servicio puede estar descargando el modelo


class VozError(Exception):
    """Error con un mensaje apto para mostrarle al usuario."""


def habilitado() -> bool:
    return bool(WHISPER_URL)


def transcribir(audio: bytes) -> str:
    if not habilitado():
        raise VozError("La búsqueda por voz no está configurada.")
    req = urllib.request.Request(f"{WHISPER_URL}/transcribe", data=audio, method="POST")
    req.add_header("Content-Type", "application/octet-stream")
    if WHISPER_TOKEN:
        req.add_header("Authorization", f"Bearer {WHISPER_TOKEN}")
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            return str(json.load(resp).get("text", "")).strip()
    except urllib.error.HTTPError as e:
        try:
            detalle = json.load(e).get("error", "")
        except Exception:
            detalle = ""
        if e.code == 401:
            raise VozError("El servicio de voz rechazó la clave (revisá WHISPER_TOKEN).")
        raise VozError(detalle or f"El servicio de voz respondió con error {e.code}.")
    except (urllib.error.URLError, TimeoutError, OSError):
        raise VozError("No pude conectar con el servicio de voz (Whisper). Probá de nuevo en un rato.")


_MULETILLAS = re.compile(
    r"^\s*(?:(?:por\s+favor|hola|che)[\s,]+)*"
    r"(?:busc[aá](?:r|me|lo)?|dame|dec[ií]me|tra[eé](?:me|r)?|mostr[aá](?:me|r)?|consult[aá](?:r|me)?|quiero\s+ver|necesito)"
    r"(?:\s+(?:a|al|el|la|los|las|de|del|info|informaci[oó]n\s+de(?:l)?))*\s+",
    re.IGNORECASE,
)


def limpiar_consulta(texto: str) -> str:
    """'Buscá a Ariel.' -> 'Ariel'. Saca el verbo inicial y la puntuación de los bordes."""
    t = (texto or "").strip()
    t = _MULETILLAS.sub("", t)
    return t.strip(" \t\n.,;:!?¿¡\"'«»")
