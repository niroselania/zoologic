# Zoologic - Consulta de Planilla

Aplicacion web independiente para cargar la planilla de Zoologic, buscar por usuario o puesto, ver la fila completa, editar datos y descargar el Excel actualizado.

## Portainer

Subir estos archivos a un repositorio Git o a una carpeta del servidor:

```text
Dockerfile
docker-compose.yml
app.py
bot.py
requirements.txt
templates/index.html
.env.example
```

El stack publica la app en el puerto `8100`:

```text
http://IP-DE-TU-SERVIDOR:8100
```

Los archivos cargados y backups quedan en el volumen `zoologic_data`.

## Uso

- Cargar el Excel desde la pantalla principal.
- Elegir la hoja.
- Buscar por usuario o puesto.
- Editar una fila desde la tabla si hace falta.
- Usar `Conectar RD` cuando la fila tenga dato en la columna `Escritorio remoto`.
- Copiar la password cuando la fila tenga dato en la columna `Password`.
- Ver dashboard automatico de filas, activas, faltantes y duplicados.
- Revisar alertas de datos incompletos o raros.
- Buscar solo por usuario/puesto o en toda la fila.
- Copiar el acceso completo de una fila para soporte.
- Consultar el historial de cambios realizados desde la app.
- Descargar el Excel actualizado.

El boton `Conectar RD` genera un archivo `.rdp`. En Windows, al abrir ese archivo se inicia la conexion con Escritorio Remoto o Windows App si esta asociada a ese formato.

El historial se guarda en `/app/data/historial_cambios.jsonl`, dentro del volumen `zoologic_data`.

## Bot de Telegram (integrado)

El stack levanta dos servicios que comparten el mismo volumen `zoologic_data`:

- `zoologic-consulta`: el sitio web.
- `zoologic-bot`: el bot de Telegram. Lee directamente `zoologic_actual.xlsx` (la planilla que se sube y edita desde el sitio), asi que siempre responde con la version actual. Si la planilla cambia, el bot lo detecta solo en la proxima consulta; no hay que subir nada aparte ni usar `/recargar`.

### Variables en Portainer (Stack > Environment variables)

| Variable | Descripcion |
|---|---|
| `TELEGRAM_BOT_TOKEN` | Token de @BotFather (obligatorio). |
| `ALLOWED_CHAT_IDS` | chat_id autorizados, separados por coma. Si queda vacio el bot rechaza a todos. |
| `SEARCH_FIELDS` | Campos donde busca (default `sector,usuario,ubicacion`). |
| `SHOW_PASSWORD` | `1` para que el bot muestre la columna Password (default oculto). |

Primer arranque: desplegar con `ALLOWED_CHAT_IDS` vacio, escribirle `/id` al bot, copiar el numero a `ALLOWED_CHAT_IDS` y redesplegar.

### Uso

Escribirle al bot un sector, usuario o ubicacion (ej: `ariel`, `remoto`, `suipacha`). Responde con la linea completa de cada coincidencia (sin distinguir mayusculas ni tildes) y la fecha de ultima modificacion de la planilla.

Comandos: `/estado` (fecha y filas de la planilla), `/cambios` (ultimos cambios hechos desde el sitio), `/id`.

Importante: usar el bot en un solo lugar. Si el bot viejo (systemd) sigue corriendo con el mismo token, Telegram da error de conflicto; hay que detenerlo (`sudo systemctl disable --now zoologic-bot`).

## Bot de Discord (opcional)

Servicio `zoologic-discord`, mismo funcionamiento que el de Telegram (lee la planilla viva del sitio).

| Variable | Descripcion |
|---|---|
| `DISCORD_BOT_TOKEN` | Token del bot (Developer Portal > Bot). Obligatorio. |
| `DISCORD_ALLOWED_USER_IDS` | IDs de usuario autorizados, separados por coma. Si queda vacio rechaza a todos. |
| `DISCORD_CHANNEL_IDS` | Opcional: limita los canales donde responde (siempre responde por mensaje directo). |

Requiere activar "Message Content Intent" en el Developer Portal. Comandos: `!id`, `!estado`, `!cambios`.
Si no usas Telegram, borra el bloque `zoologic-bot` del compose (y viceversa).

## Busqueda por voz (opcional)

Con el servicio Whisper instalado (stack aparte `whisper`), se puede mandar un mensaje de voz al bot de Telegram o de Discord y lo transcribe y busca como si fuera texto ("busca ariel").

1. Desplegar el stack `whisper` y anotar su direccion, por ejemplo `http://IP-DEL-SERVIDOR:9000`.
2. En este stack agregar `WHISPER_URL` (esa direccion) y `WHISPER_TOKEN` (la misma clave del stack whisper) y redesplegar reconstruyendo la imagen.
3. Mandar un mensaje de voz de hasta 60 segundos (`VOZ_MAX_SECONDS`). El bot responde "Entendi: ..." y los resultados.
