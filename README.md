# Control QR Asistencias - Render Free + PostgreSQL

Webapp en FastAPI para control de entradas/salidas por QR, pensada para vigilancia en móvil y monitoreo en escritorio/móvil.

Esta versión está ajustada para **Render Free**: no usa disco persistente y **no captura ni almacena imágenes**. La evidencia visual del vehículo queda fuera del sistema, por ejemplo en la bitácora de WhatsApp.

## Base de datos

Esta versión usa:

- **PostgreSQL en Render** cuando existe la variable `DATABASE_URL`.
- **SQLite local** automáticamente cuando no existe `DATABASE_URL`, para probar en tu computadora sin configurar nada.

Las tablas se crean automáticamente al iniciar la app:

- `employees`
- `attendance`
- `audit_log`
- `users`
- `user_sessions`
- `correction_batches`

## Funciones incluidas

- Registro de empleados.
- Importación masiva desde Excel.
- Edición posterior de expedientes.
- QR por ID de empleado.
- Imagen tipo credencial QR para que el empleado la muestre desde su celular.
- Impresión masiva de QRs desde navegador.
- Vista mobile first para vigilancia.
- Vista de monitoreo para escritorio y móvil.
- Registro de entrada y salida.
- Turno Día y Noche.
- Turno noche cruza medianoche.
- Registro de si el empleado entra/sale con vehículo.
- No captura placas.
- No captura fotos del personal.
- No captura fotos del vehículo en esta versión.
- Auditoría básica de altas, ediciones y registros.
- Login Admin.
- Correcciones masivas desde Excel con comparación inteligente.
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
schema_postgresql.sql
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
/login         Acceso Admin
/logout        Cerrar sesión
/vigilancia    Pantalla mobile first para entradas y salidas
/monitor       Monitoreo operativo
/personal      Lista y edición de empleados
/importar      Importación masiva desde Excel
/qr            Módulo de QR e impresión
/exportar      Exportación a Excel con filtros
/correcciones  Corrección masiva desde Excel
/healthz       Revisión rápida de app y tipo de base de datos
```

## Deploy en Render

Este proyecto incluye `render.yaml` para crear:

- Web Service de FastAPI en plan free.
- Base de datos Render Postgres en plan free.
- Variable `DATABASE_URL` conectada a Postgres.
- `DATA_DIR=/tmp/asistencias_data` solo para archivos temporales.

No usa Persistent Disk, porque Render Free no lo soporta para servicios web. Qué amable detalle de la nube, reservar la persistencia de archivos para cuando pagas.

Render usará:

```bash
pip install -r requirements.txt
```

como Build Command y:

```bash
uvicorn main:app --host 0.0.0.0 --port $PORT
```

como Start Command.

## Persistencia

- La **base de datos** queda en PostgreSQL.
- Los **registros, empleados, auditoría, correcciones y usuarios** quedan en PostgreSQL.
- Esta versión **no guarda imágenes**.
- Las imágenes QR se generan al momento desde la base de datos, no requieren disco persistente.

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

## Correcciones masivas desde Excel

Ruta:

```txt
/correcciones
```

Flujo:

```txt
Exportar Excel editable → modificar columnas permitidas → reimportar → analizar cambios → aplicar lote
```

El sistema compara el Excel contra la base de datos y aplica solo alteraciones válidas, con auditoría por campo y por lote.

Solo usuarios Admin pueden usar este módulo.

## Plantilla Excel de empleados

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
- Observaciones

## QR

El QR contiene únicamente el ID del empleado.

Ejemplo:

```txt
EMP-000123
```

El sistema usa ese ID para consultar la base de datos.

## Usuario administrador inicial

Esta versión crea automáticamente un usuario administrador inicial cuando todavía no existe ese usuario:

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
/correcciones
```

La pantalla de vigilancia (`/vigilancia`) queda disponible para operación rápida de entrada/salida.

Para producción, cambia la clave inicial desde variables de entorno o agrega después un módulo de administración de usuarios. Usar credenciales fijas eternamente es cómodo, sí, igual que dejar la llave abajo del tapete.

Variables disponibles:

```txt
DEFAULT_ADMIN_USERNAME=Admin4rd
DEFAULT_ADMIN_PASSWORD=Adm4rd
SESSION_DAYS=14
SESSION_COOKIE_SECURE=0
```

En Render puedes poner `SESSION_COOKIE_SECURE=1` cuando ya esté funcionando sobre HTTPS.
