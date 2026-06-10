# Control QR Asistencias - Render Free

Sistema web FastAPI para control de asistencias por QR con PostgreSQL en Render.

## Incluye

- Login con roles: Supremo, Admin, RH y Vigilancia.
- Usuario Supremo: `Adjm` / `Adjm4rdur`.
- Usuario Supremo adicional: `Admin4rd` / `Adm4rd`.
- Usuario Admin4rd: `Admin4rd` / `Adm4rd` con rol `Supremo`.
- Usuario Vigilancia: `Altima` / `Altima`.
- Usuario RH: `Adhm4` / `4dhm`.
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
- Vista sobria para RH.
- Importación masiva de empleados desde Excel con ID numérico automático.
- Exportación a Excel.
- Correcciones masivas desde Excel con auditoría.
- Módulo de Configuración.
- Módulo de Usuarios.
- Módulo de Estado del Sistema.
- Limpieza beta de registros de asistencia para usuario Supremo.
- Generación de QR simple por empleado.
- Generación de imagen de celular por empleado.
- Exportación ZIP de todas las imágenes de celular.
- Exportación ZIP de todos los QR simples.
- Bitácora técnica de eventos.

## Versión actual sin captura de imágenes operativas

Esta versión no captura ni almacena fotos de empleados o vehículos, para funcionar en Render Free sin disco persistente.
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
- `/rh`
- `/sistema`
- `/supremo/registros`
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

### IDs de empleados

- Los empleados nuevos usan ID únicamente numérico.
- Al crear un empleado manualmente, el sistema propone automáticamente el último ID numérico + 1.
- En importación Excel, si dejas el ID vacío, el sistema asigna automáticamente el siguiente ID disponible.
- Si capturas un ID en Excel, debe ser numérico. IDs antiguos tipo `EMP-000125` ya no se usan para nuevos empleados.


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

- Supremo: acceso total + limpieza beta de registros.
- Admin: configuración, reportes, personal, correcciones y QR.
- RH: vista sobria de asistencias y retardos.
- Vigilancia: solo `/vigilancia` y APIs necesarias de registro.

## Seguridad incluida

- Las claves no se guardan en texto plano: se almacenan con PBKDF2-SHA256 + salt.
- Cookies de sesión `HttpOnly`, `Secure` y `SameSite=Strict` en Render.
- Límite de intentos fallidos de login: 6 intentos y bloqueo temporal de 15 minutos.
- Cabeceras de seguridad: `X-Frame-Options`, `X-Content-Type-Options`, `Referrer-Policy`, `Permissions-Policy` y CSP básica.
- Roles estrictos: Vigilancia solo registra, RH solo consulta, Admin administra, Supremo limpia beta.
- Solo Supremo puede asignar/modificar usuarios Supremo.
- Validación de redirecciones internas para evitar open redirect.
- Borrado beta requiere escribir `BORRAR` y se registra en auditoría.

Nota prudente: estas claves iniciales son para beta. Para producción real conviene cambiarlas por claves largas y únicas por usuario.

## Limpieza beta

Ruta: `/supremo/registros`

Solo el rol `Supremo` puede borrar registros de asistencia para pruebas beta.
La eliminación conserva auditoría y bitácora técnica.

## Cambios de iteración: QR masivo e importación histórica

- `/qr/celular/todos.zip`: exporta todas las imágenes de celular en un ZIP optimizado. Cada PNG se nombra con nombre completo e ID.
- `/qr/simples/todos.zip`: exporta todos los QR simples en ZIP.
- `/plantilla-asistencias.xlsx`: plantilla para importar asistencias de días pasados.
- `/importar/asistencias`: carga asistencias históricas desde Excel. El sistema recalcula retardo, salida temprana y extra con base en el turno configurado.

Reglas de importación histórica:

- Requiere `ID empleado`, `Fecha turno` y `Entrada`.
- Si detecta retardo, el motivo de retardo es obligatorio.
- Si detecta salida temprana, el motivo de salida temprana es obligatorio.
- Si marcas “actualizar históricos”, modifica registros existentes que coincidan por empleado + fecha turno + turno.
- Si no marcas actualizar y ya existe un registro, no importa nada y muestra error.

## Iteración: vigilantes por QR y bajas controladas

- Nuevo módulo `/vigilantes` para crear, activar/inactivar, descargar QR y administrar códigos de cambio de turno.
- Vigilantes iniciales:
  - `VIG-MS-1` · Altima 1 · Reymundo Méndez · Activo.
  - `VIG-MS-2` · Altima 2 · David Martinez · Activo.
- En `/vigilancia`, primero se debe escanear un QR de vigilante. Ese vigilante queda como responsable activo hasta que otro vigilante escanee su QR.
- Cada entrada/salida de empleado se guarda automáticamente con el vigilante activo, no con una selección manual.
- En `/personal` se puede dar baja/reactivar trabajadores. El borrado definitivo queda solo para rol Supremo y audita la acción.
- En `/vigilantes` se puede activar/inactivar vigilantes. El borrado definitivo queda solo para rol Supremo y audita la acción.

### Prueba recomendada
1. Entrar como `Adjm`.
2. Abrir `/vigilantes` y descargar QR de `VIG-MS-1`.
3. Entrar como `Altima`.
4. Abrir `/vigilancia`.
5. Escanear el QR de vigilante.
6. Escanear un QR de empleado.
7. Verificar en `/monitor` o exportación que el registro quede ligado a `Altima 1 - Reymundo Méndez`.

## Iteración visual PC profesional

Esta versión mejora la experiencia de escritorio sin cambiar la lógica principal:

- Menú lateral profesional para roles Admin/Supremo/RH.
- Encabezado superior con usuario, rol y acceso al estado del sistema.
- Dashboard de monitoreo más ejecutivo.
- Tarjetas KPI más legibles.
- Tablas con mejor espaciado, badges y jerarquía visual.
- Módulos administrativos más sobrios en PC.
- Se conserva la experiencia Mobile First en `/vigilancia`.

La vigilancia sigue siendo simple: escanear QR de vigilante, luego QR de empleado. La administración en PC ahora se ve como sistema interno serio, no como app móvil inflada con aire triste.

## Última mejora: captura móvil y gafetes

- URL dedicada para vigilancia: `/captura`.
- La pantalla de captura móvil fue rediseñada para uso en celular: cámara protagonista, vigilante activo visible, botones grandes y paleta azul.
- El módulo `/qr` ahora permite exportar imágenes tipo celular y gafetes físicos.
- Gafetes individuales: `/qr/gafete/ID_EMPLEADO.png`.
- Exportación masiva de gafetes de 10 cm x 6 cm: `/qr/gafetes/todos.zip`.
- Los gafetes se generan como PNG a 300 dpi, con nombre completo e ID del empleado en el nombre del archivo.

## Actualización: faltas, incidencias del día e importación histórica más flexible

- El dashboard muestra un bloque de empleados activos sin registro en el día actual.
- Las incidencias recientes del dashboard ahora se filtran únicamente por la fecha actual.
- La importación histórica de asistencias detecta encabezados aunque el Excel tenga un título arriba de la tabla.
- Para cargas históricas, si falta guardia se usa `IMPORTACION_HISTORICA`.
- Para cargas históricas, si se detecta retardo o salida temprana sin motivo, se asigna `NO REGISTRADO - CARGA HISTORICA` y se agrega observación.


## Regla de ID numérico automático

Los empleados nuevos usan ID únicamente numérico. Cuando el ID se deja vacío en alta manual o importación Excel, el sistema calcula el siguiente ID tomando el número más alto de empleados operativos y sumando 1.

Para este cálculo NO se consideran empleados con:

- Área `GG`
- Puesto `GERENCIA GENERAL`

Aun así, el sistema evita duplicados: si el candidato calculado ya existe por algún registro excluido, avanza al siguiente número disponible. Porque duplicar IDs sería una forma muy eficiente de fabricar caos administrativo.

## Recálculo de reglas de asistencia

Nueva ruta:

```txt
/recalcular
```

Sirve cuando se cambian tolerancias o reglas de turno después de haber importado/capturado asistencias. El sistema recalcula por rango de fechas usando la configuración actual de turnos:

- Estado de entrada: Correcta / Retardo
- Minutos de retardo
- Estado de salida: Correcta / Salida temprana / Extra
- Minutos de salida temprana
- Minutos extra
- Límites programados y snapshots de tolerancias

No cambia las horas reales de entrada/salida. Todo cambio queda auditado como `RECALCULATE_RULES`.

Para aplicar, el Admin/Supremo debe escribir `RECALCULAR`.
