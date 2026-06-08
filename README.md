# Control QR Asistencias - Render Free

Sistema web FastAPI para control de asistencias por QR con PostgreSQL en Render.

## Incluye

- Login con roles: Admin y Vigilancia.
- Usuario inicial Admin: `Admin4rd` / `Adm4rd`.
- Pantalla de vigilancia Mobile First tipo app.
- El sistema decide automáticamente si el escaneo es entrada o salida.
- Solo puede existir una entrada abierta por empleado.
- Turnos configurables.
- Tolerancia configurable para retardo.
- Motivo obligatorio cuando hay retardo.
- Tolerancia configurable para salida temprana.
- Extra automático después de los minutos configurados.
- Cierre provisional automático de registros abiertos vencidos.
- Reporte de retardos por fecha.
- Reporte semanal acumulado.
- Importación masiva de empleados desde Excel.
- Exportación a Excel.
- Correcciones masivas desde Excel con auditoría.
- Módulo de Configuración.
- Módulo de Usuarios.
- Módulo de Estado del Sistema.
- Bitácora técnica de eventos.

## Versión actual sin imágenes

Esta versión no captura ni almacena fotos, para funcionar en Render Free sin disco persistente.
La evidencia visual del vehículo se mantiene por WhatsApp, según el flujo operativo definido.

## Rutas principales

- `/login`
- `/vigilancia`
- `/monitor`
- `/configuracion`
- `/turnos`
- `/usuarios`
- `/personal`
- `/importar`
- `/qr`
- `/retardos`
- `/reportes/semanal`
- `/exportar`
- `/correcciones`
- `/sistema`
- `/healthz`

## Despliegue en Render

El proyecto incluye `render.yaml` para crear:

- Web Service FastAPI.
- PostgreSQL Free.
- Variables de entorno necesarias.
- Python 3.13.5.

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
uvicorn main:app --host 0.0.0.0 --port $PORT
```

## Reglas operativas clave

### Entrada / salida automática

- Si no hay entrada abierta, el escaneo registra entrada.
- Si hay entrada abierta, el escaneo registra salida.
- Si la entrada abierta venció por regla de cierre provisional, se cierra automáticamente antes de permitir una nueva entrada.

### Cierre provisional

Ejemplo turno día:

- Entrada: 08:00.
- Salida programada: 19:00.
- Cierre provisional: 02:00 del día siguiente.
- Si no se registró salida antes de ese límite, el sistema cierra a las 19:00, marca `Salida provisional` y deja `Pendiente de revisión`.

### Retardos

- Se calculan según turno + tolerancia.
- Si hay retardo, el motivo es obligatorio.
- Un retardo no se borra: se conserva, se corrige o se anula con auditoría.

### Roles

- Admin: acceso total.
- Vigilancia: solo `/vigilancia` y APIs necesarias de registro.

