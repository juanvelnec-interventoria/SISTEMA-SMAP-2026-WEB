# SISTEMA SMAP 2026 - WEB

## Despliegue en Render
1. Crear un repositorio de GitHub llamado `SISTEMA-SMAP-2026-WEB`.
2. Subir el contenido de esta carpeta a la raíz del repositorio.
3. En Render: New > Web Service > conectar el repositorio.
4. Build Command: `pip install -r requirements.txt`
5. Start Command: `python server.py`
6. Crear la variable de entorno `SMAP_SYNC_TOKEN` con una clave privada.
7. Deploy.

El tablero arranca con `data.json` si existe; en este paquete se incluye una copia inicial tomada de la versión local V2.16.

## Actualización desde el PC
Después de tener la URL de Render y el token, configurar en `actualizar_smap_web.py`:
- `WEB_URL`
- `TOKEN`
- `EXCEL`

Luego ejecutar `ACTUALIZAR_SMAP_WEB.bat`. El script lee la hoja `Tabla SMAP-CIV`, conserva la lógica de `INTERVENTOR` por fila y envía los registros a `/api/sync`.

## Importante
El almacenamiento local del servicio web en Render puede ser efímero. Para producción, la siguiente etapa debe mover `data.json` a una base/almacenamiento persistente (por ejemplo Supabase) o implementar el mismo mecanismo de publicación que usa el tablero de Recorridos. El paquete está preparado para que el tablero y el actualizador estén separados.
