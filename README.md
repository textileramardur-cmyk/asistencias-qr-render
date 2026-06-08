# Control QR Asistencias - Render + PostgreSQL

Webapp en FastAPI para control de entradas/salidas por QR, pensada para vigilancia en móvil y monitoreo en escritorio/móvil.

## Base de datos

Esta versión usa:

- **PostgreSQL en Render** cuando existe la variable `DATABASE_URL`.
- **SQLite local** automáticamente cuando no existe `DATABASE_URL`, para probar en tu computadora sin configurar nada.

Las tablas se crean automáticamente al iniciar la app:

- `employees`
- `attendance`
- `audit_log`

## Funciones incluidas

- Registro de empleados.
- Importación masiva desde Excel.
- Edición posterior de expedientes.
- Foto registrada del empleado para validación visual por vigilancia.
- QR por ID de empleado.
- Imagen tipo credencial para que el empleado la muestre desde su celular.
- Impresión masiva de QRs desde navegador.
- Vista mobile first para vigilancia.
- Vista de monitoreo para escritorio y móvil.
- Registro de entrada y salida.
- Turno Día y Noche.
- Turno noche cruza medianoche.
- Captura de fotos del vehículo cuando aplica:
  - Foto frontal del vehículo.
  - Foto de cajuela.
- No captura placa como dato.
- No toma foto del personal al entrar o salir.
- Auditoría básica de altas, ediciones y registros.
- Exportación a Excel cuando se necesite:
  - empleados
  - asistencias
  - incidencias
  - auditoría
  - reporte general con varias hojas

## Estructura principal

```txt
main.py
requirements.txt
render.yaml
templates/
static/
```

## Ejecutar localmente

```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

Abrir:

```txt
http://127.0.0.1:8000
```

Localmente se creará automáticamente:

```txt
data/asistencias.db
```

## Rutas principales

```txt
/vigilancia   Pantalla mobile first para entradas y salidas
/monitor      Monitoreo operativo
/personal     Lista y edición de empleados
/importar     Importación masiva desde Excel
/qr           Módulo de QR e impresión
/exportar     Exportación a Excel con filtros
/healthz      Revisión rápida de app y tipo de base de datos
```

## Deploy en Render

Este proyecto incluye `render.yaml` para crear:

- Web Service de FastAPI.
- Base de datos Render Postgres.
- Variable `DATABASE_URL` conectada a Postgres.
- Disco persistente para fotos en `/var/data`.

Render recomienda para FastAPI:

```bash
pip install -r requirements.txt
```

como Build Command y:

```bash
uvicorn main:app --host 0.0.0.0 --port $PORT
```

como Start Command.

## Notas sobre persistencia

- La **base de datos** queda en PostgreSQL.
- Las **fotos de empleados y vehículos** se guardan en `/var/data/uploads`.
- Para que las fotos sobrevivan a redeploys/restarts, el servicio necesita el disco persistente definido en `render.yaml`.

Si Render no permite disco persistente en el plan seleccionado, cambia el plan del servicio web desde el dashboard o guarda las fotos en almacenamiento externo en una versión posterior.


## Exportación Excel

La ruta principal es:

```txt
/exportar
```

Desde ahí puedes descargar:

```txt
/exportar/general.xlsx
/exportar/asistencias.xlsx
/exportar/incidencias.xlsx
/exportar/empleados.xlsx
/exportar/auditoria.xlsx
```

Los reportes de asistencias, incidencias y general aceptan filtros por fecha de turno:

```txt
/exportar/asistencias.xlsx?fecha_inicio=2026-06-01&fecha_fin=2026-06-30
```

El reporte general incluye varias hojas: resumen, empleados, asistencias, incidencias y auditoría.

## Plantilla Excel

La plantilla se descarga desde:

```txt
/plantilla-empleados.xlsx
```

Columnas esperadas:

- ID empleado
- Nombre completo
- Area
- Puesto
- Turno
- Estado
- Tiene vehiculo
- Requiere fotos vehiculo
- Observaciones

## QR

El QR contiene únicamente el ID del empleado.

Ejemplo:

```txt
EMP-000123
```

El sistema usa ese ID para consultar la base de datos.

## Fotos

Regla de diseño:

- No se toma foto del personal en entrada ni salida.
- Solo se muestra la foto registrada del empleado.
- Si el empleado tiene vehículo, se capturan fotos del vehículo.
- No se captura placa como campo.

## Usuario administrador inicial

Esta versión crea automáticamente un usuario administrador inicial cuando la base de datos está vacía o cuando todavía no existe ese usuario:

```txt
Usuario: Admin4rd
Clave: Adm4rd
Rol: Admin
```

Rutas administrativas protegidas por login:

```txt
/login
/logout
/monitor
/personal
/importar
/qr
/exportar
```

La pantalla de vigilancia (`/vigilancia`) queda disponible para operación rápida de entrada/salida. Las altas, ediciones, importación, exportación, QR y monitoreo quedan detrás del acceso Admin.

Para producción, cambia la clave inicial desde variables de entorno o agrega después un módulo de administración de usuarios. Usar credenciales fijas eternamente es cómodo, sí, igual que dejar la llave abajo del tapete.

Variables disponibles:

```txt
DEFAULT_ADMIN_USERNAME=Admin4rd
DEFAULT_ADMIN_PASSWORD=Adm4rd
SESSION_DAYS=14
SESSION_COOKIE_SECURE=0
```

En Render puedes poner `SESSION_COOKIE_SECURE=1` cuando ya esté funcionando sobre HTTPS.

## Correcciones masivas por Excel

Ruta principal:

```txt
/correcciones
```

Flujo:

1. Admin descarga un Excel editable desde `/correcciones`.
2. Modifica únicamente columnas nuevas, por ejemplo:
   - `entrada_nueva`
   - `salida_nueva`
   - `turno_nuevo`
   - `fecha_turno_nueva`
   - `motivo_retardo_nuevo`
   - `motivo_salida_temprana_nuevo`
   - `observaciones_nuevas`
   - `motivo_correccion`
3. Reimporta el Excel.
4. El sistema compara contra la base actual.
5. Detecta cambios por campo.
6. Valida errores y control de versión (`updated_at_actual`).
7. Muestra previsualización.
8. Solo Admin puede aplicar el lote.
9. Cada cambio queda registrado en auditoría con folio de lote.

Reglas:

- Celda vacía en columna nueva = no modificar.
- Para borrar un valor editable, escribir `BORRAR`.
- Si falta `motivo_correccion`, el lote no se aplica.
- Si el registro cambió después de exportar, se bloquea esa fila y se debe exportar de nuevo.
