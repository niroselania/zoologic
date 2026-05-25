# Zoologic - Consulta de Planilla

Aplicacion web independiente para cargar la planilla de Zoologic, buscar por usuario o puesto, ver la fila completa, editar datos y descargar el Excel actualizado.

## Portainer

Subir estos archivos a un repositorio Git o a una carpeta del servidor:

```text
Dockerfile
docker-compose.yml
app.py
requirements.txt
templates/index.html
.dockerignore
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
- Descargar el Excel actualizado.
