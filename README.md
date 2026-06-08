# Control QR Asistencias - Render Free

Sistema web FastAPI para control de asistencias mediante QR, preparado para Render Free con PostgreSQL.

## Versión actual

- Sin almacenamiento de imágenes/fotos en el sistema.
- Evidencia visual de vehículo por bitácora externa de WhatsApp.
- QR contiene únicamente el ID del empleado.
- PostgreSQL en Render mediante `DATABASE_URL`.
- Usuario administrador inicial:
  - Usuario: `Admin4rd`
  - Clave: `Adm4rd`

## Funciones principales

- `/vigilancia`: captura QR Mobile First estilo cámara, registro de entrada/salida.
- `/monitor`: monitoreo operativo.
- `/retardos`: vista rápida por fecha de retardos, salidas tempranas y extras.
- `/turnos`: configuración de turnos, tolerancias y regla de extra.
- `/personal`: alta y edición de empleados.
- `/importar`: importación masiva desde Excel.
- `/qr`: generación e impresión de QRs.
- `/exportar`: exportación a Excel.
- `/correcciones`: corrección masiva tipo exportar Excel editable, modificar, reimportar, analizar y aplicar.

## Reglas operativas implementadas

### Retardo

El sistema calcula retardo automáticamente:

`entrada real > entrada programada + tolerancia configurable`

Si hay retardo, el motivo es obligatorio.

### Salida temprana

`salida real < salida programada - tolerancia configurable`

Si hay salida temprana, el motivo es obligatorio.

### Extra

`salida real > salida programada + minutos configurados para extra`

Por defecto, Extra se considera después de 30 minutos de la salida programada.

## Deploy en Render

El repositorio debe tener estos archivos en la raíz:

```txt
main.py
requirements.txt
render.yaml
schema_postgresql.sql
README.md
templates/
static/
```

Render debe crear:

- Web Service `asistencias-qr-render`
- PostgreSQL `asistencias-qr-db`

Prueba de salud:

```txt
/healthz
```

Debe responder:

```json
{"ok": true, "db": "postgresql"}
```
