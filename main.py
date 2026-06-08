import hashlib
import io
import json
import os
import re
import secrets
from datetime import datetime, date, time, timedelta
from urllib.parse import quote
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

import qrcode
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Connection

APP_NAME = "Control QR Asistencias"
TZ = ZoneInfo("America/Mexico_City")
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.getenv("DATA_DIR", BASE_DIR / "data"))
UPLOADS_DIR = DATA_DIR / "uploads"
CORRECTIONS_DIR = DATA_DIR / "correcciones"
EMPLOYEE_PHOTOS_DIR = UPLOADS_DIR / "empleados"
VEHICLE_PHOTOS_DIR = UPLOADS_DIR / "vehiculos"
DB_PATH = DATA_DIR / "asistencias.db"

AUTH_COOKIE_NAME = "asistencias_session"
SESSION_DAYS = int(os.getenv("SESSION_DAYS", "14"))
SESSION_COOKIE_SECURE = os.getenv("SESSION_COOKIE_SECURE", "0") == "1"
DEFAULT_ADMIN_USERNAME = os.getenv("DEFAULT_ADMIN_USERNAME", "Admin4rd")
DEFAULT_ADMIN_PASSWORD = os.getenv("DEFAULT_ADMIN_PASSWORD", "Adm4rd")

DATA_DIR.mkdir(parents=True, exist_ok=True)
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
CORRECTIONS_DIR.mkdir(parents=True, exist_ok=True)
EMPLOYEE_PHOTOS_DIR.mkdir(parents=True, exist_ok=True)
VEHICLE_PHOTOS_DIR.mkdir(parents=True, exist_ok=True)

raw_database_url = os.getenv("DATABASE_URL", "").strip()
if raw_database_url:
    # Render puede entregar postgresql://; SQLAlchemy también lo acepta en versiones nuevas,
    # pero normalizamos a psycopg para evitar dramas de dependencias, que ya tenemos suficientes.
    if raw_database_url.startswith("postgres://"):
        DATABASE_URL = raw_database_url.replace("postgres://", "postgresql+psycopg://", 1)
    elif raw_database_url.startswith("postgresql://"):
        DATABASE_URL = raw_database_url.replace("postgresql://", "postgresql+psycopg://", 1)
    else:
        DATABASE_URL = raw_database_url
else:
    DATABASE_URL = f"sqlite:///{DB_PATH}"

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, future=True, pool_pre_ping=True, connect_args=connect_args)

app = FastAPI(title=APP_NAME)
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
app.mount("/uploads", StaticFiles(directory=str(UPLOADS_DIR)), name="uploads")


# -----------------------------
# Helpers
# -----------------------------

def now_mx() -> datetime:
    return datetime.now(TZ)


def clean_id(value: str) -> str:
    value = (value or "").strip().upper()
    value = re.sub(r"[^A-Z0-9_\-]", "", value)
    return value[:40]


def safe_filename(value: str) -> str:
    value = clean_id(value) or "ARCHIVO"
    return value


def bool_from_excel(value) -> int:
    if value is None:
        return 0
    text_value = str(value).strip().lower()
    return 1 if text_value in {"si", "sí", "s", "yes", "y", "true", "1", "x"} else 0


def as_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def row_to_dict(row):
    if row is None:
        return None
    return dict(row)


def fetch_one(conn: Connection, sql: str, params: Optional[dict] = None):
    row = conn.execute(text(sql), params or {}).mappings().first()
    return dict(row) if row else None


def fetch_all(conn: Connection, sql: str, params: Optional[dict] = None):
    rows = conn.execute(text(sql), params or {}).mappings().all()
    return [dict(r) for r in rows]


def init_db() -> None:
    is_sqlite = engine.dialect.name == "sqlite"
    attendance_id = "INTEGER PRIMARY KEY AUTOINCREMENT" if is_sqlite else "SERIAL PRIMARY KEY"
    audit_id = "INTEGER PRIMARY KEY AUTOINCREMENT" if is_sqlite else "SERIAL PRIMARY KEY"

    statements = [
        """
        CREATE TABLE IF NOT EXISTS employees (
            id TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            area TEXT DEFAULT '',
            puesto TEXT DEFAULT '',
            turno TEXT DEFAULT 'Día',
            estado TEXT DEFAULT 'Activo',
            tiene_vehiculo INTEGER DEFAULT 0,
            requiere_fotos_vehiculo INTEGER DEFAULT 0,
            foto_path TEXT DEFAULT '',
            qr_activo INTEGER DEFAULT 1,
            observaciones TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        f"""
        CREATE TABLE IF NOT EXISTS attendance (
            id {attendance_id},
            employee_id TEXT NOT NULL,
            shift_date TEXT NOT NULL,
            turno TEXT NOT NULL,
            entry_at TEXT,
            exit_at TEXT,
            entry_guard TEXT DEFAULT '',
            exit_guard TEXT DEFAULT '',
            entry_status TEXT DEFAULT '',
            exit_status TEXT DEFAULT '',
            late_reason TEXT DEFAULT '',
            early_reason TEXT DEFAULT '',
            vehicle_expected INTEGER DEFAULT 0,
            vehicle_entered INTEGER DEFAULT 0,
            vehicle_front_entry TEXT DEFAULT '',
            vehicle_trunk_entry TEXT DEFAULT '',
            vehicle_front_exit TEXT DEFAULT '',
            vehicle_trunk_exit TEXT DEFAULT '',
            incident TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password_hash TEXT NOT NULL,
            role TEXT DEFAULT 'Admin',
            active INTEGER DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS user_sessions (
            token_hash TEXT PRIMARY KEY,
            username TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(username) REFERENCES users(username)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS correction_batches (
            batch_id TEXT PRIMARY KEY,
            user_name TEXT NOT NULL,
            status TEXT DEFAULT 'preview',
            file_path TEXT DEFAULT '',
            total_rows INTEGER DEFAULT 0,
            total_changes INTEGER DEFAULT 0,
            total_errors INTEGER DEFAULT 0,
            created_at TEXT NOT NULL,
            applied_at TEXT DEFAULT ''
        )
        """,
        f"""
        CREATE TABLE IF NOT EXISTS audit_log (
            id {audit_id},
            table_name TEXT NOT NULL,
            record_id TEXT NOT NULL,
            action TEXT NOT NULL,
            user_name TEXT DEFAULT '',
            field_name TEXT DEFAULT '',
            old_value TEXT DEFAULT '',
            new_value TEXT DEFAULT '',
            reason TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )
        """,
        "CREATE INDEX IF NOT EXISTS idx_attendance_employee_open ON attendance(employee_id, exit_at)",
        "CREATE INDEX IF NOT EXISTS idx_attendance_shift_date ON attendance(shift_date)",
        "CREATE INDEX IF NOT EXISTS idx_audit_record ON audit_log(table_name, record_id)",
        "CREATE INDEX IF NOT EXISTS idx_sessions_username ON user_sessions(username)",
        "CREATE INDEX IF NOT EXISTS idx_sessions_expires ON user_sessions(expires_at)",
    ]

    with engine.begin() as conn:
        for statement in statements:
            conn.exec_driver_sql(statement)


def audit(conn: Connection, table: str, record_id: str, action: str,
          user: str = "Sistema", field: str = "", old: str = "", new: str = "", reason: str = "") -> None:
    conn.execute(
        text(
            """
            INSERT INTO audit_log (table_name, record_id, action, user_name, field_name, old_value, new_value, reason, created_at)
            VALUES (:table_name, :record_id, :action, :user_name, :field_name, :old_value, :new_value, :reason, :created_at)
            """
        ),
        {
            "table_name": table,
            "record_id": record_id,
            "action": action,
            "user_name": user,
            "field_name": field,
            "old_value": old,
            "new_value": new,
            "reason": reason,
            "created_at": now_mx().isoformat(),
        }
    )


def get_employee(conn: Connection, employee_id: str):
    return fetch_one(conn, "SELECT * FROM employees WHERE id = :id", {"id": employee_id})


def current_shift_date(turno: str, dt: Optional[datetime] = None) -> str:
    dt = dt or now_mx()
    # Si el turno es noche y estamos antes de medio día, normalmente pertenece al día anterior.
    if (turno or "").lower().startswith("n") and dt.time() < time(12, 0):
        return (dt.date() - timedelta(days=1)).isoformat()
    return dt.date().isoformat()


def expected_times(turno: str, shift_date_iso: str):
    shift_day = date.fromisoformat(shift_date_iso)
    if (turno or "").lower().startswith("n"):
        start_dt = datetime.combine(shift_day, time(19, 0), tzinfo=TZ)
        end_dt = datetime.combine(shift_day + timedelta(days=1), time(8, 0), tzinfo=TZ)
    else:
        start_dt = datetime.combine(shift_day, time(8, 0), tzinfo=TZ)
        end_dt = datetime.combine(shift_day, time(19, 0), tzinfo=TZ)
    return start_dt, end_dt


def entry_status(turno: str, shift_date_iso: str, dt: datetime, tolerance_minutes: int = 10) -> str:
    start_dt, _ = expected_times(turno, shift_date_iso)
    if dt > start_dt + timedelta(minutes=tolerance_minutes):
        return "Tarde"
    return "Correcta"


def exit_status(turno: str, shift_date_iso: str, dt: datetime, tolerance_minutes: int = 10) -> str:
    _, end_dt = expected_times(turno, shift_date_iso)
    if dt < end_dt - timedelta(minutes=tolerance_minutes):
        return "Salida temprana"
    return "Correcta"


async def save_upload(file: Optional[UploadFile], folder: Path, prefix: str) -> str:
    if not file or not file.filename:
        return ""
    suffix = Path(file.filename).suffix.lower() or ".jpg"
    if suffix not in {".jpg", ".jpeg", ".png", ".webp"}:
        suffix = ".jpg"
    name = f"{safe_filename(prefix)}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{suffix}"
    path = folder / name
    content = await file.read()
    path.write_bytes(content)
    return str(path.relative_to(UPLOADS_DIR)).replace("\\", "/")


# -----------------------------
# Usuarios, login y sesiones
# -----------------------------

def hash_password(password: str, salt: Optional[str] = None) -> str:
    salt = salt or secrets.token_hex(16)
    iterations = 200_000
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), iterations)
    return f"pbkdf2_sha256${iterations}${salt}${digest.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        algorithm, iterations_text, salt, expected = stored_hash.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), int(iterations_text)).hex()
        return secrets.compare_digest(digest, expected)
    except Exception:
        return False


def session_hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def parse_dt(value: str) -> Optional[datetime]:
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=TZ)
        return dt.astimezone(TZ)
    except Exception:
        return None


def seed_admin_user() -> None:
    with engine.begin() as conn:
        existing = fetch_one(conn, "SELECT * FROM users WHERE username = :username", {"username": DEFAULT_ADMIN_USERNAME})
        if existing:
            return
        ts = now_mx().isoformat()
        conn.execute(
            text(
                """
                INSERT INTO users (username, password_hash, role, active, created_at, updated_at)
                VALUES (:username, :password_hash, 'Admin', 1, :created_at, :updated_at)
                """
            ),
            {
                "username": DEFAULT_ADMIN_USERNAME,
                "password_hash": hash_password(DEFAULT_ADMIN_PASSWORD),
                "created_at": ts,
                "updated_at": ts,
            }
        )
        audit(conn, "users", DEFAULT_ADMIN_USERNAME, "CREATE", "Sistema", reason="Usuario administrador inicial")


def get_current_user(request: Request):
    token = request.cookies.get(AUTH_COOKIE_NAME, "")
    if not token:
        return None
    token_digest = session_hash(token)
    with engine.begin() as conn:
        row = fetch_one(
            conn,
            """
            SELECT s.token_hash, s.username, s.expires_at, u.role, u.active
            FROM user_sessions s
            JOIN users u ON u.username = s.username
            WHERE s.token_hash = :token_hash
            """,
            {"token_hash": token_digest},
        )
        if not row or not row.get("active"):
            return None
        expires = parse_dt(row.get("expires_at", ""))
        if not expires or expires < now_mx():
            conn.execute(text("DELETE FROM user_sessions WHERE token_hash = :token_hash"), {"token_hash": token_digest})
            return None
        return {"username": row["username"], "role": row["role"]}


def admin_login_redirect(request: Request) -> RedirectResponse:
    target = request.url.path
    if request.url.query:
        target += "?" + request.url.query
    return RedirectResponse(url=f"/login?next={quote(target)}", status_code=303)


def require_admin_page(request: Request):
    user = get_current_user(request)
    if not user:
        return None
    if user.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="Se requiere usuario Admin")
    return user


def require_admin_http(request: Request):
    user = require_admin_page(request)
    if not user:
        raise HTTPException(status_code=401, detail="Inicia sesión como Admin")
    return user


templates.env.globals["get_current_user"] = get_current_user


# -----------------------------
# Excel export helpers
# -----------------------------

def bool_text(value) -> str:
    return "Sí" if bool(value) else "No"


def short_datetime(value) -> str:
    if not value:
        return ""
    text_value = str(value)
    try:
        dt = datetime.fromisoformat(text_value.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return text_value[:16]


def date_filter_clause(alias: str = "a", fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None):
    conditions = []
    params = {}
    if fecha_inicio:
        conditions.append(f"{alias}.shift_date >= :fecha_inicio")
        params["fecha_inicio"] = fecha_inicio
    if fecha_fin:
        conditions.append(f"{alias}.shift_date <= :fecha_fin")
        params["fecha_fin"] = fecha_fin
    return (" AND " + " AND ".join(conditions) if conditions else "", params)


def style_worksheet(ws, title: str = "") -> None:
    header_fill = PatternFill("solid", fgColor="051A39")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(color="051A39", bold=True, size=16)
    thin = Side(style="thin", color="D9DEE5")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    if title:
        ws.insert_rows(1)
        ws["A1"] = title
        ws["A1"].font = title_font
        ws["A1"].alignment = Alignment(horizontal="left")
        ws.freeze_panes = "A3"
        header_row = 2
    else:
        ws.freeze_panes = "A2"
        header_row = 1

    max_row = ws.max_row
    max_col = ws.max_column
    if max_row < header_row:
        return

    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row in ws.iter_rows(min_row=header_row + 1, max_row=max_row, max_col=max_col):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border

    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col)}{max_row}"

    for col_idx in range(1, max_col + 1):
        letter = get_column_letter(col_idx)
        max_len = 10
        for cell in ws[letter]:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(len(value), 42))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 38)


def workbook_response(wb: Workbook, filename: str):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


def make_sheet(wb: Workbook, name: str, headers: list[str], rows: list[list], title: str = ""):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_worksheet(ws, title or name)
    return ws


def remove_default_sheet(wb: Workbook) -> None:
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]


def get_employees_export_rows(conn: Connection):
    employees = fetch_all(conn, "SELECT * FROM employees ORDER BY nombre")
    headers = [
        "ID empleado", "Nombre", "Área", "Puesto", "Turno", "Estado", "Tiene vehículo",
        "QR activo", "Observaciones", "Creado", "Actualizado"
    ]
    rows = [
        [
            emp["id"], emp["nombre"], emp["area"], emp["puesto"], emp["turno"], emp["estado"],
            bool_text(emp["tiene_vehiculo"]), bool_text(emp["qr_activo"]),
            emp["observaciones"], short_datetime(emp["created_at"]), short_datetime(emp["updated_at"]),
        ]
        for emp in employees
    ]
    return headers, rows


def get_attendance_export_rows(conn: Connection, fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None):
    clause, params = date_filter_clause("a", fecha_inicio, fecha_fin)
    records = fetch_all(conn, f"""
        SELECT a.*, e.nombre, e.area, e.puesto
        FROM attendance a
        LEFT JOIN employees e ON e.id = a.employee_id
        WHERE 1=1 {clause}
        ORDER BY a.shift_date DESC, a.entry_at DESC
    """, params)
    headers = [
        "ID registro", "ID empleado", "Nombre", "Área", "Puesto", "Fecha turno", "Turno",
        "Entrada", "Salida", "Guardia entrada", "Guardia salida", "Estado entrada", "Estado salida",
        "Motivo retardo", "Motivo salida temprana", "Vehículo esperado", "Vehículo registrado",
        "Incidencia / observaciones", "Creado", "Actualizado"
    ]
    rows = [
        [
            row["id"], row["employee_id"], row.get("nombre") or "", row.get("area") or "", row.get("puesto") or "",
            row["shift_date"], row["turno"], short_datetime(row["entry_at"]), short_datetime(row["exit_at"]),
            row["entry_guard"], row["exit_guard"], row["entry_status"], row["exit_status"], row["late_reason"],
            row["early_reason"], bool_text(row["vehicle_expected"]), bool_text(row["vehicle_entered"]),
            row["incident"], short_datetime(row["created_at"]), short_datetime(row["updated_at"]),
        ]
        for row in records
    ]
    return headers, rows


def get_incidents_export_rows(conn: Connection, fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None):
    clause, params = date_filter_clause("a", fecha_inicio, fecha_fin)
    records = fetch_all(conn, f"""
        SELECT a.*, e.nombre, e.area, e.puesto
        FROM attendance a
        LEFT JOIN employees e ON e.id = a.employee_id
        WHERE 1=1 {clause}
          AND (
            COALESCE(a.entry_status, '') = 'Tarde'
            OR COALESCE(a.exit_status, '') = 'Salida temprana'
            OR COALESCE(a.incident, '') != ''
          )
        ORDER BY a.updated_at DESC
    """, params)
    headers = [
        "ID registro", "ID empleado", "Nombre", "Área", "Fecha turno", "Turno", "Tipo incidencia",
        "Entrada", "Salida", "Motivo", "Observaciones", "Vehículo registrado"
    ]
    rows = []
    for row in records:
        tipo = []
        if row["entry_status"] == "Tarde":
            tipo.append("Retardo")
        if row["exit_status"] == "Salida temprana":
            tipo.append("Salida temprana")
        if row["incident"]:
            tipo.append("Observación")
        rows.append([
            row["id"], row["employee_id"], row.get("nombre") or "", row.get("area") or "", row["shift_date"], row["turno"],
            ", ".join(tipo) or "Incidencia", short_datetime(row["entry_at"]), short_datetime(row["exit_at"]),
            row["late_reason"] or row["early_reason"] or "", row["incident"], bool_text(row["vehicle_entered"]),
        ])
    return headers, rows


def get_audit_export_rows(conn: Connection):
    records = fetch_all(conn, "SELECT * FROM audit_log ORDER BY created_at DESC")
    headers = ["ID", "Tabla", "ID registro", "Acción", "Usuario", "Campo", "Valor anterior", "Valor nuevo", "Motivo", "Fecha"]
    rows = [
        [
            row["id"], row["table_name"], row["record_id"], row["action"], row["user_name"], row["field_name"],
            row["old_value"], row["new_value"], row["reason"], short_datetime(row["created_at"]),
        ]
        for row in records
    ]
    return headers, rows


def add_summary_sheet(wb: Workbook, conn: Connection, fecha_inicio: Optional[str], fecha_fin: Optional[str]) -> None:
    ws = wb.create_sheet("Resumen")
    ws.append(["Reporte general de asistencia"])
    ws.append(["Generado", now_mx().strftime("%Y-%m-%d %H:%M")])
    ws.append(["Fecha inicio", fecha_inicio or "Sin filtro"])
    ws.append(["Fecha fin", fecha_fin or "Sin filtro"])
    ws.append([])

    clause, params = date_filter_clause("a", fecha_inicio, fecha_fin)
    total_empleados = conn.execute(text("SELECT COUNT(*) FROM employees")).scalar() or 0
    activos = conn.execute(text("SELECT COUNT(*) FROM employees WHERE estado = 'Activo'")).scalar() or 0
    total_asistencias = conn.execute(text(f"SELECT COUNT(*) FROM attendance a WHERE 1=1 {clause}"), params).scalar() or 0
    retardos = conn.execute(text(f"SELECT COUNT(*) FROM attendance a WHERE entry_status = 'Tarde' {clause}"), params).scalar() or 0
    salidas_tempranas = conn.execute(text(f"SELECT COUNT(*) FROM attendance a WHERE exit_status = 'Salida temprana' {clause}"), params).scalar() or 0
    vehiculos = conn.execute(text(f"SELECT COUNT(*) FROM attendance a WHERE vehicle_entered = 1 {clause}"), params).scalar() or 0
    abiertas = conn.execute(text(f"SELECT COUNT(*) FROM attendance a WHERE exit_at IS NULL {clause}"), params).scalar() or 0

    ws.append(["Indicador", "Valor"])
    for label, value in [
        ("Empleados registrados", total_empleados),
        ("Empleados activos", activos),
        ("Registros de asistencia", total_asistencias),
        ("Retardos", retardos),
        ("Salidas tempranas", salidas_tempranas),
        ("Entradas con vehículo", vehiculos),
        ("Entradas abiertas", abiertas),
    ]:
        ws.append([label, value])

    ws.append([])
    ws.append(["Nota", "El QR contiene únicamente el ID del empleado. Esta versión no captura imágenes ni placas; la evidencia visual se mantiene por WhatsApp."])
    style_worksheet(ws, "")
    ws["A1"].font = Font(color="051A39", bold=True, size=16)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 42



# -----------------------------
# Correcciones masivas desde Excel
# -----------------------------

def normalize_header(value) -> str:
    txt = as_text(value).lower()
    replacements = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ñ": "n",
        "_": " ", "-": " ", "/": " ", ".": " ",
    }
    for old, new in replacements.items():
        txt = txt.replace(old, new)
    return " ".join(txt.split())


def parse_date_value(value) -> Optional[str]:
    if value is None or as_text(value) == "":
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    txt = as_text(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt, fmt).date().isoformat()
        except Exception:
            pass
    return txt


def parse_time_or_datetime(value, shift_date_iso: str, turno: str, kind: str, entry_dt: Optional[datetime] = None):
    if value is None or as_text(value) == "":
        return None, ""
    if as_text(value).upper() == "BORRAR":
        return "__CLEAR__", ""
    try:
        if isinstance(value, datetime):
            dt = value
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=TZ)
            return dt.astimezone(TZ), ""
        if isinstance(value, time):
            t = value
        elif isinstance(value, (int, float)):
            # Excel puede guardar una hora como fracción del día: 0.5 = 12:00. Gracias, Excel, muy normal todo.
            if 0 <= float(value) < 1:
                total_seconds = int(round(float(value) * 24 * 60 * 60))
                t = (datetime.min + timedelta(seconds=total_seconds)).time()
            else:
                return None, f"Valor numérico de hora inválido: {value}"
        else:
            txt = as_text(value)
            # Si viene fecha y hora completa, usarla tal cual.
            for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S"):
                try:
                    dt = datetime.strptime(txt, fmt).replace(tzinfo=TZ)
                    return dt, ""
                except Exception:
                    pass
            # Si solo viene hora, combinar con fecha de turno.
            for fmt in ("%H:%M", "%H:%M:%S"):
                try:
                    t = datetime.strptime(txt, fmt).time()
                    break
                except Exception:
                    t = None
            if t is None:
                return None, f"Formato de hora inválido: {txt}"

        base_date = date.fromisoformat(shift_date_iso)
        if kind == "salida" and (turno or "").lower().startswith("n"):
            base_date = base_date + timedelta(days=1)
        dt = datetime.combine(base_date, t, tzinfo=TZ)
        if kind == "salida" and entry_dt and dt < entry_dt:
            dt = dt + timedelta(days=1)
        return dt, ""
    except Exception as exc:
        return None, f"No se pudo interpretar fecha/hora: {exc}"


def correction_export_headers() -> list[str]:
    return [
        "id_registro", "id_empleado", "nombre", "area", "fecha_turno_actual", "turno_actual",
        "entrada_actual", "salida_actual", "estado_entrada_actual", "estado_salida_actual",
        "motivo_retardo_actual", "motivo_salida_temprana_actual", "observaciones_actuales", "updated_at_actual",
        "entrada_nueva", "salida_nueva", "turno_nuevo", "fecha_turno_nueva",
        "motivo_retardo_nuevo", "motivo_salida_temprana_nuevo", "observaciones_nuevas", "motivo_correccion",
    ]


def build_editable_corrections_workbook(conn: Connection, fecha_inicio: Optional[str], fecha_fin: Optional[str]) -> Workbook:
    clause, params = date_filter_clause("a", fecha_inicio, fecha_fin)
    records = fetch_all(conn, f"""
        SELECT a.*, e.nombre, e.area
        FROM attendance a
        LEFT JOIN employees e ON e.id = a.employee_id
        WHERE 1=1 {clause}
        ORDER BY a.shift_date DESC, a.entry_at DESC
    """, params)
    wb = Workbook()
    ws = wb.active
    ws.title = "correcciones"
    headers = correction_export_headers()
    ws.append(headers)
    for row in records:
        ws.append([
            row["id"], row["employee_id"], row.get("nombre") or "", row.get("area") or "",
            row["shift_date"], row["turno"], short_datetime(row["entry_at"]), short_datetime(row["exit_at"]),
            row["entry_status"], row["exit_status"], row["late_reason"], row["early_reason"], row["incident"], row["updated_at"],
            "", "", "", "", "", "", "", "",
        ])
    style_worksheet(ws, "Corrección masiva editable")
    note = wb.create_sheet("instrucciones")
    note.append(["Uso"])
    note.append(["Edita solo columnas que terminan en _nueva y motivo_correccion."])
    note.append(["Una celda nueva vacía significa: no modificar."])
    note.append(["Para borrar un valor editable, escribe BORRAR."])
    note.append(["Al reimportar, el sistema compara contra la base, detecta alteraciones, valida y audita cada cambio."])
    style_worksheet(note, "Instrucciones")
    return wb


def analyze_corrections_excel(content: bytes, admin_username: str):
    errors = []
    changes = []
    total_rows = 0
    try:
        wb = load_workbook(io.BytesIO(content), data_only=True)
    except Exception as exc:
        return {"ok": False, "errors": [f"No se pudo leer el Excel: {exc}"], "changes": [], "total_rows": 0}
    ws = wb["correcciones"] if "correcciones" in wb.sheetnames else wb.active
    raw_headers = [cell.value for cell in ws[1]]
    headers = {normalize_header(h): idx for idx, h in enumerate(raw_headers)}

    def col(name: str):
        return headers.get(normalize_header(name))

    required = ["id_registro", "id_empleado", "updated_at_actual"]
    missing = [name for name in required if col(name) is None]
    if missing:
        return {"ok": False, "errors": ["Faltan columnas requeridas: " + ", ".join(missing)], "changes": [], "total_rows": 0}

    seen_records = set()
    with engine.begin() as conn:
        for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            total_rows += 1
            record_id = as_text(row[col("id_registro")])
            employee_id = clean_id(as_text(row[col("id_empleado")]))
            version_excel = as_text(row[col("updated_at_actual")])
            motivo = as_text(row[col("motivo_correccion")]) if col("motivo_correccion") is not None else ""
            if not record_id:
                errors.append(f"Fila {row_number}: falta id_registro")
                continue
            if record_id in seen_records:
                errors.append(f"Fila {row_number}: registro duplicado en el Excel: {record_id}")
                continue
            seen_records.add(record_id)
            db = fetch_one(conn, "SELECT * FROM attendance WHERE id = :id", {"id": record_id})
            if not db:
                errors.append(f"Fila {row_number}: registro {record_id} no existe")
                continue
            if clean_id(db["employee_id"]) != employee_id:
                errors.append(f"Fila {row_number}: el empleado no coincide con el registro {record_id}")
                continue
            if as_text(db["updated_at"]) != version_excel:
                errors.append(f"Fila {row_number}: el registro {record_id} fue modificado después de exportar. Exporta de nuevo esa fila.")
                continue

            new_shift_date = db["shift_date"]
            if col("fecha_turno_nueva") is not None:
                val = row[col("fecha_turno_nueva")]
                if as_text(val):
                    if as_text(val).upper() == "BORRAR":
                        errors.append(f"Fila {row_number}: fecha_turno_nueva no se puede borrar")
                        continue
                    parsed = parse_date_value(val)
                    if not parsed:
                        errors.append(f"Fila {row_number}: fecha_turno_nueva inválida")
                        continue
                    new_shift_date = parsed

            new_turno = db["turno"]
            if col("turno_nuevo") is not None:
                val = as_text(row[col("turno_nuevo")])
                if val:
                    if val.lower() in {"dia", "día"}:
                        new_turno = "Día"
                    elif val.lower() == "noche":
                        new_turno = "Noche"
                    else:
                        errors.append(f"Fila {row_number}: turno_nuevo inválido: {val}")
                        continue

            row_changes = []
            def add_change(field, old, new):
                old_txt = "" if old is None else str(old)
                new_txt = "" if new is None else str(new)
                if old_txt != new_txt:
                    row_changes.append({"record_id": record_id, "field": field, "old": old_txt, "new": new_txt, "row_number": row_number, "reason": motivo})

            # Horas
            entry_dt_current = parse_dt(db["entry_at"]) if db.get("entry_at") else None
            entry_value = db.get("entry_at") or ""
            if col("entrada_nueva") is not None and as_text(row[col("entrada_nueva")]):
                parsed, err = parse_time_or_datetime(row[col("entrada_nueva")], new_shift_date, new_turno, "entrada")
                if err:
                    errors.append(f"Fila {row_number}: {err}")
                    continue
                entry_value = "" if parsed == "__CLEAR__" else parsed.isoformat()
                entry_dt_current = None if parsed == "__CLEAR__" else parsed
                add_change("entry_at", db.get("entry_at") or "", entry_value)

            exit_value = db.get("exit_at") or ""
            if col("salida_nueva") is not None and as_text(row[col("salida_nueva")]):
                parsed, err = parse_time_or_datetime(row[col("salida_nueva")], new_shift_date, new_turno, "salida", entry_dt_current)
                if err:
                    errors.append(f"Fila {row_number}: {err}")
                    continue
                exit_value = "" if parsed == "__CLEAR__" else parsed.isoformat()
                add_change("exit_at", db.get("exit_at") or "", exit_value)

            if entry_value and exit_value:
                entry_dt = parse_dt(entry_value)
                exit_dt = parse_dt(exit_value)
                if entry_dt and exit_dt and exit_dt < entry_dt:
                    errors.append(f"Fila {row_number}: la salida queda antes que la entrada")
                    continue

            add_change("shift_date", db["shift_date"], new_shift_date)
            add_change("turno", db["turno"], new_turno)

            # Textos editables
            field_map = [
                ("motivo_retardo_nuevo", "late_reason"),
                ("motivo_salida_temprana_nuevo", "early_reason"),
                ("observaciones_nuevas", "incident"),
            ]
            for excel_col, db_field in field_map:
                if col(excel_col) is not None and as_text(row[col(excel_col)]):
                    val = as_text(row[col(excel_col)])
                    new_val = "" if val.upper() == "BORRAR" else val
                    add_change(db_field, db.get(db_field) or "", new_val)

            # Recalcular estados si se movieron horas/turno/fecha.
            entry_for_status = parse_dt(entry_value) if entry_value else None
            exit_for_status = parse_dt(exit_value) if exit_value else None
            if entry_for_status:
                add_change("entry_status", db.get("entry_status") or "", entry_status(new_turno, new_shift_date, entry_for_status))
            if exit_for_status:
                add_change("exit_status", db.get("exit_status") or "", exit_status(new_turno, new_shift_date, exit_for_status))

            if row_changes and not motivo:
                errors.append(f"Fila {row_number}: falta motivo_correccion")
                continue
            changes.extend(row_changes)

    return {"ok": len(errors) == 0, "errors": errors, "changes": changes, "total_rows": total_rows}


def save_correction_preview(analysis: dict, admin_username: str, original_filename: str = "") -> str:
    batch_id = f"CORR-{now_mx().strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(3).upper()}"
    path = CORRECTIONS_DIR / f"{batch_id}.json"
    payload = {"batch_id": batch_id, "admin": admin_username, "original_filename": original_filename, "analysis": analysis, "created_at": now_mx().isoformat()}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    with engine.begin() as conn:
        conn.execute(text("""
            INSERT INTO correction_batches (batch_id, user_name, status, file_path, total_rows, total_changes, total_errors, created_at)
            VALUES (:batch_id, :user_name, 'preview', :file_path, :total_rows, :total_changes, :total_errors, :created_at)
        """), {
            "batch_id": batch_id,
            "user_name": admin_username,
            "file_path": str(path),
            "total_rows": analysis.get("total_rows", 0),
            "total_changes": len(analysis.get("changes", [])),
            "total_errors": len(analysis.get("errors", [])),
            "created_at": now_mx().isoformat(),
        })
    return batch_id


def apply_correction_batch(batch_id: str, admin_username: str) -> dict:
    batch_id = clean_id(batch_id)
    path = CORRECTIONS_DIR / f"{batch_id}.json"
    if not path.exists():
        return {"ok": False, "message": "Lote no encontrado"}
    payload = json.loads(path.read_text(encoding="utf-8"))
    analysis = payload.get("analysis", {})
    if analysis.get("errors"):
        return {"ok": False, "message": "No se puede aplicar un lote con errores"}
    changes = analysis.get("changes", [])
    if not changes:
        return {"ok": False, "message": "No hay cambios para aplicar"}

    grouped = {}
    for change in changes:
        grouped.setdefault(str(change["record_id"]), []).append(change)

    allowed_fields = {"entry_at", "exit_at", "shift_date", "turno", "late_reason", "early_reason", "incident", "entry_status", "exit_status"}
    applied = 0
    ts = now_mx().isoformat()
    with engine.begin() as conn:
        batch = fetch_one(conn, "SELECT * FROM correction_batches WHERE batch_id = :batch_id", {"batch_id": batch_id})
        if batch and batch.get("status") == "applied":
            return {"ok": False, "message": "Este lote ya fue aplicado"}
        for record_id, record_changes in grouped.items():
            set_parts = []
            params = {"id": record_id, "updated_at": ts}
            for change in record_changes:
                field = change["field"]
                if field not in allowed_fields:
                    continue
                param_name = f"v_{field}"
                set_parts.append(f"{field}=:{param_name}")
                if field in {"entry_at", "exit_at"} and change["new"] == "":
                    params[param_name] = None
                else:
                    params[param_name] = change["new"]
            if not set_parts:
                continue
            set_parts.append("updated_at=:updated_at")
            conn.execute(text(f"UPDATE attendance SET {', '.join(set_parts)} WHERE id=:id"), params)
            for change in record_changes:
                audit(conn, "attendance", record_id, "BULK_UPDATE", admin_username, field=change["field"], old=change["old"], new=change["new"], reason=f"{batch_id}: {change.get('reason','')}")
                applied += 1
        conn.execute(text("UPDATE correction_batches SET status='applied', applied_at=:applied_at WHERE batch_id=:batch_id"), {"applied_at": ts, "batch_id": batch_id})
    return {"ok": True, "message": f"Lote aplicado. Cambios: {applied}", "applied": applied}

def seed_demo_if_empty() -> None:
    with engine.begin() as conn:
        count = conn.execute(text("SELECT COUNT(*) FROM employees")).scalar() or 0
        if count:
            return
        ts = now_mx().isoformat()
        demo = [
            {
                "id": "EMP-000123", "nombre": "Juan Pérez", "area": "Producción", "puesto": "Operador", "turno": "Día",
                "estado": "Activo", "tiene_vehiculo": 1, "requiere_fotos_vehiculo": 0, "foto_path": "", "qr_activo": 1,
                "observaciones": "Demo con vehículo", "created_at": ts, "updated_at": ts,
            },
            {
                "id": "EMP-000124", "nombre": "María López", "area": "Calidad", "puesto": "Inspectora", "turno": "Día",
                "estado": "Activo", "tiene_vehiculo": 0, "requiere_fotos_vehiculo": 0, "foto_path": "", "qr_activo": 1,
                "observaciones": "Demo sin vehículo", "created_at": ts, "updated_at": ts,
            },
            {
                "id": "EMP-000125", "nombre": "Carlos Ramos", "area": "Almacén", "puesto": "Auxiliar", "turno": "Noche",
                "estado": "Activo", "tiene_vehiculo": 1, "requiere_fotos_vehiculo": 0, "foto_path": "", "qr_activo": 1,
                "observaciones": "Demo turno noche", "created_at": ts, "updated_at": ts,
            },
        ]
        for item in demo:
            conn.execute(
                text(
                    """
                    INSERT INTO employees (id, nombre, area, puesto, turno, estado, tiene_vehiculo, requiere_fotos_vehiculo, foto_path, qr_activo, observaciones, created_at, updated_at)
                    VALUES (:id, :nombre, :area, :puesto, :turno, :estado, :tiene_vehiculo, :requiere_fotos_vehiculo, :foto_path, :qr_activo, :observaciones, :created_at, :updated_at)
                    """
                ), item
            )


@app.on_event("startup")
def startup() -> None:
    init_db()
    seed_admin_user()
    seed_demo_if_empty()

# Inicialización defensiva para pruebas locales y algunos runners.
init_db()
seed_admin_user()
seed_demo_if_empty()


# -----------------------------
# Pages
# -----------------------------

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request, next: str = "/monitor", error: str = ""):
    return templates.TemplateResponse("login.html", {"request": request, "next": next, "error": error})


@app.post("/login")
def login_submit(request: Request, username: str = Form(...), password: str = Form(...), next: str = Form("/monitor")):
    with engine.begin() as conn:
        user = fetch_one(conn, "SELECT * FROM users WHERE username = :username AND active = 1", {"username": username.strip()})
        if not user or not verify_password(password, user["password_hash"]):
            return templates.TemplateResponse(
                "login.html",
                {"request": request, "next": next or "/monitor", "error": "Usuario o clave incorrectos."},
                status_code=401,
            )
        token = secrets.token_urlsafe(32)
        expires_at = (now_mx() + timedelta(days=SESSION_DAYS)).isoformat()
        conn.execute(
            text(
                """
                INSERT INTO user_sessions (token_hash, username, expires_at, created_at)
                VALUES (:token_hash, :username, :expires_at, :created_at)
                """
            ),
            {"token_hash": session_hash(token), "username": user["username"], "expires_at": expires_at, "created_at": now_mx().isoformat()},
        )
        audit(conn, "users", user["username"], "LOGIN", user["username"], reason="Inicio de sesión")

    safe_next = next if next and next.startswith("/") else "/monitor"
    response = RedirectResponse(url=safe_next, status_code=303)
    response.set_cookie(
        AUTH_COOKIE_NAME,
        token,
        httponly=True,
        secure=SESSION_COOKIE_SECURE,
        samesite="lax",
        max_age=SESSION_DAYS * 24 * 60 * 60,
    )
    return response


@app.get("/logout")
def logout(request: Request):
    token = request.cookies.get(AUTH_COOKIE_NAME, "")
    if token:
        with engine.begin() as conn:
            user = get_current_user(request)
            conn.execute(text("DELETE FROM user_sessions WHERE token_hash = :token_hash"), {"token_hash": session_hash(token)})
            if user:
                audit(conn, "users", user["username"], "LOGOUT", user["username"], reason="Cierre de sesión")
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie(AUTH_COOKIE_NAME)
    return response


@app.get("/", response_class=HTMLResponse)
def home():
    return RedirectResponse(url="/vigilancia")


@app.get("/vigilancia", response_class=HTMLResponse)
def vigilancia(request: Request):
    return templates.TemplateResponse("vigilancia.html", {"request": request})


@app.get("/monitor", response_class=HTMLResponse)
def monitor(request: Request):
    admin_user = require_admin_page(request)
    if not admin_user:
        return admin_login_redirect(request)
    today = now_mx().date().isoformat()
    with engine.begin() as conn:
        total_employees = conn.execute(text("SELECT COUNT(*) FROM employees WHERE estado = 'Activo'")).scalar() or 0
        present_today = conn.execute(
            text("SELECT COUNT(*) FROM attendance WHERE shift_date = :today AND entry_at IS NOT NULL"),
            {"today": today}
        ).scalar() or 0
        late_today = conn.execute(
            text("SELECT COUNT(*) FROM attendance WHERE shift_date = :today AND entry_status = 'Tarde'"),
            {"today": today}
        ).scalar() or 0
        vehicles_inside = conn.execute(
            text("SELECT COUNT(*) FROM attendance WHERE exit_at IS NULL AND vehicle_entered = 1")
        ).scalar() or 0
        inside = fetch_all(conn,
            """
            SELECT a.*, e.nombre, e.area, e.puesto
            FROM attendance a
            JOIN employees e ON e.id = a.employee_id
            WHERE a.exit_at IS NULL
            ORDER BY a.entry_at DESC
            """
        )
        incidents = fetch_all(conn,
            """
            SELECT a.*, e.nombre, e.area
            FROM attendance a
            JOIN employees e ON e.id = a.employee_id
            WHERE a.entry_status = 'Tarde'
               OR a.exit_status = 'Salida temprana'
               OR a.incident != ''
            ORDER BY a.updated_at DESC
            LIMIT 40
            """
        )
    return templates.TemplateResponse(
        "monitor.html",
        {
            "request": request,
            "total_employees": total_employees,
            "present_today": present_today,
            "late_today": late_today,
            "vehicles_inside": vehicles_inside,
            "inside": inside,
            "incidents": incidents,
        }
    )


@app.get("/personal", response_class=HTMLResponse)
def empleados(request: Request):
    admin_user = require_admin_page(request)
    if not admin_user:
        return admin_login_redirect(request)
    with engine.begin() as conn:
        rows = fetch_all(conn, "SELECT * FROM employees ORDER BY nombre")
    return templates.TemplateResponse("personal.html", {"request": request, "employees": rows})


@app.get("/personal/nuevo", response_class=HTMLResponse)
def empleado_nuevo(request: Request):
    admin_user = require_admin_page(request)
    if not admin_user:
        return admin_login_redirect(request)
    return templates.TemplateResponse("empleado_form.html", {"request": request, "employee": None})


@app.get("/personal/{employee_id}", response_class=HTMLResponse)
def empleado_editar(request: Request, employee_id: str):
    admin_user = require_admin_page(request)
    if not admin_user:
        return admin_login_redirect(request)
    admin_user = require_admin_http(request)
    employee_id = clean_id(employee_id)
    with engine.begin() as conn:
        emp = get_employee(conn, employee_id)
        if not emp:
            raise HTTPException(status_code=404, detail="Empleado no encontrado")
    return templates.TemplateResponse("empleado_form.html", {"request": request, "employee": emp})


@app.post("/personal/guardar")
async def empleado_guardar(
    request: Request,
    id: str = Form(...),
    nombre: str = Form(...),
    area: str = Form(""),
    puesto: str = Form(""),
    turno: str = Form("Día"),
    estado: str = Form("Activo"),
    tiene_vehiculo: str = Form("0"),
    requiere_fotos_vehiculo: str = Form("0"),
    qr_activo: str = Form("1"),
    observaciones: str = Form(""),
    foto: Optional[UploadFile] = File(None),
):
    admin_user = require_admin_http(request)
    employee_id = clean_id(id)
    if not employee_id or not nombre.strip():
        raise HTTPException(status_code=400, detail="ID y nombre son obligatorios")
    ts = now_mx().isoformat()
    foto_path = ""
    with engine.begin() as conn:
        exists = get_employee(conn, employee_id)
        if exists:
            raise HTTPException(status_code=400, detail="Ese ID ya existe. Edita el expediente existente.")
        conn.execute(
            text(
                """
                INSERT INTO employees (id, nombre, area, puesto, turno, estado, tiene_vehiculo, requiere_fotos_vehiculo, foto_path, qr_activo, observaciones, created_at, updated_at)
                VALUES (:id, :nombre, :area, :puesto, :turno, :estado, :tiene_vehiculo, :requiere_fotos_vehiculo, :foto_path, :qr_activo, :observaciones, :created_at, :updated_at)
                """
            ),
            {
                "id": employee_id,
                "nombre": nombre.strip(),
                "area": area.strip(),
                "puesto": puesto.strip(),
                "turno": turno,
                "estado": estado,
                "tiene_vehiculo": 1 if tiene_vehiculo == "1" else 0,
                "requiere_fotos_vehiculo": 1 if requiere_fotos_vehiculo == "1" else 0,
                "foto_path": foto_path,
                "qr_activo": 1 if qr_activo == "1" else 0,
                "observaciones": observaciones.strip(),
                "created_at": ts,
                "updated_at": ts,
            }
        )
        audit(conn, "employees", employee_id, "CREATE", admin_user["username"])
    return RedirectResponse(url="/personal", status_code=303)


@app.post("/personal/{employee_id}/guardar")
async def empleado_actualizar(
    request: Request,
    employee_id: str,
    nombre: str = Form(...),
    area: str = Form(""),
    puesto: str = Form(""),
    turno: str = Form("Día"),
    estado: str = Form("Activo"),
    tiene_vehiculo: str = Form("0"),
    requiere_fotos_vehiculo: str = Form("0"),
    qr_activo: str = Form("1"),
    observaciones: str = Form(""),
    foto: Optional[UploadFile] = File(None),
):
    admin_user = require_admin_http(request)
    employee_id = clean_id(employee_id)
    with engine.begin() as conn:
        emp = get_employee(conn, employee_id)
        if not emp:
            raise HTTPException(status_code=404, detail="Empleado no encontrado")
        foto_path = emp["foto_path"] or ""
        # Versión sin imágenes: no se cargan ni reemplazan fotos desde expediente.
        conn.execute(
            text(
                """
                UPDATE employees
                SET nombre=:nombre, area=:area, puesto=:puesto, turno=:turno, estado=:estado,
                    tiene_vehiculo=:tiene_vehiculo, requiere_fotos_vehiculo=:requiere_fotos_vehiculo,
                    foto_path=:foto_path, qr_activo=:qr_activo, observaciones=:observaciones, updated_at=:updated_at
                WHERE id=:id
                """
            ),
            {
                "nombre": nombre.strip(),
                "area": area.strip(),
                "puesto": puesto.strip(),
                "turno": turno,
                "estado": estado,
                "tiene_vehiculo": 1 if tiene_vehiculo == "1" else 0,
                "requiere_fotos_vehiculo": 1 if requiere_fotos_vehiculo == "1" else 0,
                "foto_path": foto_path,
                "qr_activo": 1 if qr_activo == "1" else 0,
                "observaciones": observaciones.strip(),
                "updated_at": now_mx().isoformat(),
                "id": employee_id,
            }
        )
        audit(conn, "employees", employee_id, "UPDATE", admin_user["username"], reason="Edición manual de expediente")
    return RedirectResponse(url="/personal", status_code=303)


@app.get("/importar", response_class=HTMLResponse)
def importar_page(request: Request):
    admin_user = require_admin_page(request)
    if not admin_user:
        return admin_login_redirect(request)
    return templates.TemplateResponse("importar.html", {"request": request, "result": None})


@app.get("/plantilla-empleados.xlsx")
def plantilla_empleados(request: Request):
    require_admin_http(request)
    wb = Workbook()
    ws = wb.active
    ws.title = "empleados"
    headers = [
        "ID empleado",
        "Nombre completo",
        "Area",
        "Puesto",
        "Turno",
        "Estado",
        "Tiene vehiculo",
        "Observaciones",
    ]
    ws.append(headers)
    ws.append(["EMP-000126", "Nombre Apellido", "Producción", "Operador", "Día", "Activo", "Sí", "Ejemplo"])
    ws.append(["EMP-000127", "Nombre Apellido", "Calidad", "Inspectora", "Noche", "Activo", "No", "Ejemplo"])
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[chr(64 + col)].width = 24
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=plantilla-empleados.xlsx"},
    )


@app.post("/importar", response_class=HTMLResponse)
async def importar_empleados(request: Request, archivo: UploadFile = File(...), actualizar: Optional[str] = Form(None)):
    admin_user = require_admin_http(request)
    content = await archivo.read()
    try:
        wb = load_workbook(io.BytesIO(content), data_only=True)
    except Exception as exc:
        return templates.TemplateResponse("importar.html", {"request": request, "result": {"ok": False, "errors": [f"No se pudo leer el Excel: {exc}"]}})

    ws = wb.active
    raw_headers = [as_text(cell.value) for cell in ws[1]]
    normalized = {h.lower().strip(): idx for idx, h in enumerate(raw_headers)}

    aliases = {
        "id": ["id empleado", "id", "empleado", "codigo", "código"],
        "nombre": ["nombre completo", "nombre"],
        "area": ["area", "área"],
        "puesto": ["puesto"],
        "turno": ["turno"],
        "estado": ["estado"],
        "tiene_vehiculo": ["tiene vehiculo", "tiene vehículo", "vehiculo", "vehículo"],
        "requiere_fotos_vehiculo": ["requiere fotos vehiculo", "requiere fotos vehículo", "requiere foto vehiculo", "requiere foto vehículo"],
        "observaciones": ["observaciones", "obs"],
    }

    def index_for(key: str) -> Optional[int]:
        for alias in aliases[key]:
            if alias in normalized:
                return normalized[alias]
        return None

    required = ["id", "nombre", "turno", "estado"]
    missing = [key for key in required if index_for(key) is None]
    if missing:
        return templates.TemplateResponse("importar.html", {"request": request, "result": {"ok": False, "errors": ["Faltan columnas requeridas: " + ", ".join(missing)]}})

    errors = []
    rows_to_import = []
    seen = set()

    for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(row):
            continue
        employee_id = clean_id(as_text(row[index_for("id")]))
        nombre = as_text(row[index_for("nombre")])
        turno = as_text(row[index_for("turno")]) or "Día"
        estado = as_text(row[index_for("estado")]) or "Activo"
        area = as_text(row[index_for("area")]) if index_for("area") is not None else ""
        puesto = as_text(row[index_for("puesto")]) if index_for("puesto") is not None else ""
        observaciones = as_text(row[index_for("observaciones")]) if index_for("observaciones") is not None else ""
        tiene_vehiculo = bool_from_excel(row[index_for("tiene_vehiculo")]) if index_for("tiene_vehiculo") is not None else 0
        requiere_fotos = 0

        if not employee_id:
            errors.append(f"Fila {row_number}: falta ID empleado")
        if not nombre:
            errors.append(f"Fila {row_number}: falta nombre")
        if turno not in {"Día", "Dia", "Noche"}:
            errors.append(f"Fila {row_number}: turno inválido: {turno}. Usa Día o Noche")
        if estado not in {"Activo", "Inactivo", "Baja", "Suspendido"}:
            errors.append(f"Fila {row_number}: estado inválido: {estado}")
        if employee_id in seen:
            errors.append(f"Fila {row_number}: ID duplicado dentro del Excel: {employee_id}")
        seen.add(employee_id)
        if turno == "Dia":
            turno = "Día"
        rows_to_import.append((employee_id, nombre, area, puesto, turno, estado, tiene_vehiculo, requiere_fotos, observaciones))

    with engine.begin() as conn:
        for employee_id, *_ in rows_to_import:
            exists = get_employee(conn, employee_id)
            if exists and not actualizar:
                errors.append(f"ID {employee_id}: ya existe. Marca 'actualizar existentes' si quieres modificarlo.")

    if errors:
        return templates.TemplateResponse("importar.html", {"request": request, "result": {"ok": False, "errors": errors[:200], "total_errors": len(errors)}})

    nuevos = 0
    actualizados = 0
    ts = now_mx().isoformat()
    with engine.begin() as conn:
        for employee_id, nombre, area, puesto, turno, estado, tiene_vehiculo, requiere_fotos, observaciones in rows_to_import:
            exists = get_employee(conn, employee_id)
            if exists:
                conn.execute(
                    text(
                        """
                        UPDATE employees
                        SET nombre=:nombre, area=:area, puesto=:puesto, turno=:turno, estado=:estado,
                            tiene_vehiculo=:tiene_vehiculo, requiere_fotos_vehiculo=:requiere_fotos_vehiculo,
                            observaciones=:observaciones, updated_at=:updated_at
                        WHERE id=:id
                        """
                    ),
                    {
                        "id": employee_id,
                        "nombre": nombre,
                        "area": area,
                        "puesto": puesto,
                        "turno": turno,
                        "estado": estado,
                        "tiene_vehiculo": tiene_vehiculo,
                        "requiere_fotos_vehiculo": requiere_fotos,
                        "observaciones": observaciones,
                        "updated_at": ts,
                    }
                )
                audit(conn, "employees", employee_id, "UPDATE", admin_user["username"], reason="Importación Excel")
                actualizados += 1
            else:
                conn.execute(
                    text(
                        """
                        INSERT INTO employees (id, nombre, area, puesto, turno, estado, tiene_vehiculo, requiere_fotos_vehiculo, foto_path, qr_activo, observaciones, created_at, updated_at)
                        VALUES (:id, :nombre, :area, :puesto, :turno, :estado, :tiene_vehiculo, :requiere_fotos_vehiculo, '', 1, :observaciones, :created_at, :updated_at)
                        """
                    ),
                    {
                        "id": employee_id,
                        "nombre": nombre,
                        "area": area,
                        "puesto": puesto,
                        "turno": turno,
                        "estado": estado,
                        "tiene_vehiculo": tiene_vehiculo,
                        "requiere_fotos_vehiculo": requiere_fotos,
                        "observaciones": observaciones,
                        "created_at": ts,
                        "updated_at": ts,
                    }
                )
                audit(conn, "employees", employee_id, "CREATE", admin_user["username"], reason="Importación Excel")
                nuevos += 1

    return templates.TemplateResponse("importar.html", {"request": request, "result": {"ok": True, "leidos": len(rows_to_import), "nuevos": nuevos, "actualizados": actualizados}})



@app.get("/correcciones", response_class=HTMLResponse)
def correcciones_page(request: Request):
    admin_user = require_admin_page(request)
    if not admin_user:
        return admin_login_redirect(request)
    today = now_mx().date()
    month_start = today.replace(day=1).isoformat()
    with engine.begin() as conn:
        batches = fetch_all(conn, "SELECT * FROM correction_batches ORDER BY created_at DESC LIMIT 20")
    return templates.TemplateResponse(
        "correcciones.html",
        {
            "request": request,
            "fecha_inicio": month_start,
            "fecha_fin": today.isoformat(),
            "result": None,
            "batches": batches,
        },
    )


@app.get("/correcciones/exportar-editable.xlsx")
def correcciones_exportar_editable(request: Request, fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None):
    require_admin_http(request)
    with engine.begin() as conn:
        wb = build_editable_corrections_workbook(conn, fecha_inicio, fecha_fin)
    filename = f"correcciones_editables_{fecha_inicio or 'todo'}_{fecha_fin or 'todo'}_{now_mx().strftime('%Y%m%d_%H%M')}.xlsx"
    return workbook_response(wb, filename)


@app.post("/correcciones/analizar", response_class=HTMLResponse)
async def correcciones_analizar(request: Request, archivo: UploadFile = File(...)):
    admin_user = require_admin_http(request)
    content = await archivo.read()
    analysis = analyze_corrections_excel(content, admin_user["username"])
    batch_id = save_correction_preview(analysis, admin_user["username"], archivo.filename or "")
    today = now_mx().date()
    month_start = today.replace(day=1).isoformat()
    with engine.begin() as conn:
        batches = fetch_all(conn, "SELECT * FROM correction_batches ORDER BY created_at DESC LIMIT 20")
    return templates.TemplateResponse(
        "correcciones.html",
        {
            "request": request,
            "fecha_inicio": month_start,
            "fecha_fin": today.isoformat(),
            "result": {"batch_id": batch_id, "analysis": analysis},
            "batches": batches,
        },
    )


@app.post("/correcciones/aplicar", response_class=HTMLResponse)
def correcciones_aplicar(request: Request, batch_id: str = Form(...)):
    admin_user = require_admin_http(request)
    apply_result = apply_correction_batch(batch_id, admin_user["username"])
    today = now_mx().date()
    month_start = today.replace(day=1).isoformat()
    with engine.begin() as conn:
        batches = fetch_all(conn, "SELECT * FROM correction_batches ORDER BY created_at DESC LIMIT 20")
    return templates.TemplateResponse(
        "correcciones.html",
        {
            "request": request,
            "fecha_inicio": month_start,
            "fecha_fin": today.isoformat(),
            "result": {"apply_result": apply_result},
            "batches": batches,
        },
    )


@app.get("/qr", response_class=HTMLResponse)
def qr_page(request: Request):
    admin_user = require_admin_page(request)
    if not admin_user:
        return admin_login_redirect(request)
    with engine.begin() as conn:
        employees = fetch_all(conn, "SELECT * FROM employees ORDER BY nombre")
    return templates.TemplateResponse("qr.html", {"request": request, "employees": employees})




@app.get("/exportar", response_class=HTMLResponse)
def exportar_page(request: Request):
    admin_user = require_admin_page(request)
    if not admin_user:
        return admin_login_redirect(request)
    today = now_mx().date()
    month_start = today.replace(day=1).isoformat()
    return templates.TemplateResponse(
        "exportar.html",
        {
            "request": request,
            "fecha_inicio": month_start,
            "fecha_fin": today.isoformat(),
        },
    )


@app.get("/exportar/empleados.xlsx")
def exportar_empleados_excel(request: Request):
    require_admin_http(request)
    wb = Workbook()
    with engine.begin() as conn:
        headers, rows = get_employees_export_rows(conn)
    make_sheet(wb, "Empleados", headers, rows, "Empleados")
    remove_default_sheet(wb)
    filename = f"empleados_{now_mx().strftime('%Y%m%d_%H%M')}.xlsx"
    return workbook_response(wb, filename)


@app.get("/exportar/asistencias.xlsx")
def exportar_asistencias_excel(request: Request, fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None):
    require_admin_http(request)
    wb = Workbook()
    title = "Asistencias"
    if fecha_inicio or fecha_fin:
        title += f" ({fecha_inicio or 'inicio'} a {fecha_fin or 'fin'})"
    with engine.begin() as conn:
        headers, rows = get_attendance_export_rows(conn, fecha_inicio, fecha_fin)
    make_sheet(wb, "Asistencias", headers, rows, title)
    remove_default_sheet(wb)
    filename = f"asistencias_{fecha_inicio or 'todo'}_{fecha_fin or 'todo'}_{now_mx().strftime('%Y%m%d_%H%M')}.xlsx"
    return workbook_response(wb, filename)


@app.get("/exportar/incidencias.xlsx")
def exportar_incidencias_excel(request: Request, fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None):
    require_admin_http(request)
    wb = Workbook()
    title = "Incidencias"
    if fecha_inicio or fecha_fin:
        title += f" ({fecha_inicio or 'inicio'} a {fecha_fin or 'fin'})"
    with engine.begin() as conn:
        headers, rows = get_incidents_export_rows(conn, fecha_inicio, fecha_fin)
    make_sheet(wb, "Incidencias", headers, rows, title)
    remove_default_sheet(wb)
    filename = f"incidencias_{fecha_inicio or 'todo'}_{fecha_fin or 'todo'}_{now_mx().strftime('%Y%m%d_%H%M')}.xlsx"
    return workbook_response(wb, filename)


@app.get("/exportar/auditoria.xlsx")
def exportar_auditoria_excel(request: Request):
    require_admin_http(request)
    wb = Workbook()
    with engine.begin() as conn:
        headers, rows = get_audit_export_rows(conn)
    make_sheet(wb, "Auditoria", headers, rows, "Auditoría")
    remove_default_sheet(wb)
    filename = f"auditoria_{now_mx().strftime('%Y%m%d_%H%M')}.xlsx"
    return workbook_response(wb, filename)


@app.get("/exportar/general.xlsx")
def exportar_general_excel(request: Request, fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None):
    require_admin_http(request)
    wb = Workbook()
    with engine.begin() as conn:
        add_summary_sheet(wb, conn, fecha_inicio, fecha_fin)
        emp_headers, emp_rows = get_employees_export_rows(conn)
        att_headers, att_rows = get_attendance_export_rows(conn, fecha_inicio, fecha_fin)
        inc_headers, inc_rows = get_incidents_export_rows(conn, fecha_inicio, fecha_fin)
        aud_headers, aud_rows = get_audit_export_rows(conn)
    make_sheet(wb, "Empleados", emp_headers, emp_rows, "Empleados")
    make_sheet(wb, "Asistencias", att_headers, att_rows, "Asistencias")
    make_sheet(wb, "Incidencias", inc_headers, inc_rows, "Incidencias")
    make_sheet(wb, "Auditoria", aud_headers, aud_rows, "Auditoría")
    remove_default_sheet(wb)
    filename = f"reporte_general_{fecha_inicio or 'todo'}_{fecha_fin or 'todo'}_{now_mx().strftime('%Y%m%d_%H%M')}.xlsx"
    return workbook_response(wb, filename)

@app.get("/qr/{employee_id}.png")
def qr_png(employee_id: str):
    employee_id = clean_id(employee_id)
    with engine.begin() as conn:
        emp = get_employee(conn, employee_id)
        if not emp:
            raise HTTPException(status_code=404, detail="Empleado no encontrado")
    img = qrcode.make(employee_id).convert("RGB")
    img = img.resize((900, 900))
    output = io.BytesIO()
    img.save(output, format="PNG")
    output.seek(0)
    return StreamingResponse(output, media_type="image/png", headers={"Content-Disposition": f"inline; filename={employee_id}.png"})


def load_font(size: int, bold: bool = False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def center_text(draw: ImageDraw.ImageDraw, xy, text_value: str, font, fill):
    x, y, width = xy
    bbox = draw.textbbox((0, 0), text_value, font=font)
    tw = bbox[2] - bbox[0]
    draw.text((x + (width - tw) / 2, y), text_value, font=font, fill=fill)


@app.get("/qr/card/{employee_id}.png")
def qr_card_png(employee_id: str):
    employee_id = clean_id(employee_id)
    with engine.begin() as conn:
        emp = get_employee(conn, employee_id)
        if not emp:
            raise HTTPException(status_code=404, detail="Empleado no encontrado")

    W, H = 1080, 1920
    azul = (5, 26, 57)
    verde = (36, 198, 166)
    gris = (245, 247, 250)
    blanco = (255, 255, 255)
    oscuro = (24, 32, 42)

    img = Image.new("RGB", (W, H), gris)
    draw = ImageDraw.Draw(img)
    draw.rectangle((0, 0, W, 310), fill=azul)
    draw.rounded_rectangle((70, 230, W - 70, H - 120), radius=54, fill=blanco)

    font_title = load_font(58, True)
    font_name = load_font(68, True)
    font_label = load_font(32, False)
    font_id = load_font(44, True)
    font_small = load_font(28, False)

    center_text(draw, (0, 74, W), "CREDENCIAL QR", font_title, blanco)
    center_text(draw, (0, 150, W), "Control de asistencia", font_label, (220, 235, 245))

    qr_img = qrcode.make(employee_id).convert("RGB").resize((720, 720))
    img.paste(qr_img, ((W - 720) // 2, 440))

    center_text(draw, (100, 1230, W - 200), emp["nombre"], font_name, oscuro)
    center_text(draw, (100, 1330, W - 200), employee_id, font_id, azul)

    info = f"{emp['area'] or 'Área no asignada'} · {emp['turno'] or 'Turno no asignado'}"
    center_text(draw, (100, 1410, W - 200), info, font_label, (80, 90, 105))

    vehicle_text = "Con vehículo" if emp["tiene_vehiculo"] else "Sin vehículo"
    badge_w, badge_h = 420, 72
    bx = (W - badge_w) // 2
    by = 1510
    draw.rounded_rectangle((bx, by, bx + badge_w, by + badge_h), radius=36, fill=verde)
    center_text(draw, (bx, by + 17, badge_w), vehicle_text, font_label, azul)

    center_text(draw, (100, 1710, W - 200), "Muestra este QR a vigilancia", font_small, (100, 110, 125))
    center_text(draw, (100, 1760, W - 200), "El QR contiene únicamente el ID del empleado", font_small, (100, 110, 125))

    output = io.BytesIO()
    img.save(output, format="PNG")
    output.seek(0)
    return StreamingResponse(output, media_type="image/png", headers={"Content-Disposition": f"inline; filename=credencial-{employee_id}.png"})


# -----------------------------
# API para vigilancia
# -----------------------------

@app.get("/api/empleado/{employee_id}")
def api_empleado(employee_id: str):
    employee_id = clean_id(employee_id)
    with engine.begin() as conn:
        emp = get_employee(conn, employee_id)
        if not emp:
            return JSONResponse(status_code=404, content={"ok": False, "message": "Empleado no encontrado"})
        open_att = fetch_one(
            conn,
            "SELECT * FROM attendance WHERE employee_id = :employee_id AND exit_at IS NULL ORDER BY id DESC LIMIT 1",
            {"employee_id": employee_id}
        )
    return {
        "ok": True,
        "employee": emp,
        "foto_url": "",
        "has_open_attendance": bool(open_att),
        "open_attendance": open_att,
    }


@app.post("/api/registro")
async def api_registro(
    employee_id: str = Form(...),
    movimiento: str = Form(...),
    guardia: str = Form("Vigilancia"),
    vehiculo: str = Form("0"),
    motivo_retardo: str = Form(""),
    motivo_salida_temprana: str = Form(""),
    observaciones: str = Form(""),
    foto_frontal: Optional[UploadFile] = File(None),
    foto_cajuela: Optional[UploadFile] = File(None),
):
    employee_id = clean_id(employee_id)
    dt = now_mx()
    ts = dt.isoformat()

    with engine.begin() as conn:
        emp = get_employee(conn, employee_id)
        if not emp:
            return JSONResponse(status_code=404, content={"ok": False, "message": "Empleado no encontrado"})
        if emp["estado"] != "Activo":
            return JSONResponse(status_code=400, content={"ok": False, "message": f"Empleado no activo: {emp['estado']}"})
        if not emp["qr_activo"]:
            return JSONResponse(status_code=400, content={"ok": False, "message": "QR inactivo"})

    vehicle_required = bool(int(vehiculo or "0")) or bool(emp["tiene_vehiculo"])

    # Versión Render Free: no se capturan ni almacenan imágenes.
    # La evidencia visual del vehículo se mantiene en bitácora externa de WhatsApp.
    front_path = ""
    trunk_path = ""

    with engine.begin() as conn:
        if movimiento == "entrada":
            open_att = fetch_one(
                conn,
                "SELECT * FROM attendance WHERE employee_id = :employee_id AND exit_at IS NULL ORDER BY id DESC LIMIT 1",
                {"employee_id": employee_id}
            )
            if open_att:
                return JSONResponse(status_code=400, content={"ok": False, "message": "Ya existe una entrada abierta para este empleado"})

            shift_date = current_shift_date(emp["turno"], dt)
            status = entry_status(emp["turno"], shift_date, dt)
            if status == "Tarde" and not motivo_retardo.strip():
                return JSONResponse(status_code=400, content={"ok": False, "message": "Entrada tarde: captura motivo de retardo"})

            result = conn.execute(
                text(
                    """
                    INSERT INTO attendance (
                        employee_id, shift_date, turno, entry_at, entry_guard, entry_status, late_reason,
                        vehicle_expected, vehicle_entered, vehicle_front_entry, vehicle_trunk_entry, incident,
                        created_at, updated_at
                    ) VALUES (
                        :employee_id, :shift_date, :turno, :entry_at, :entry_guard, :entry_status, :late_reason,
                        :vehicle_expected, :vehicle_entered, :vehicle_front_entry, :vehicle_trunk_entry, :incident,
                        :created_at, :updated_at
                    ) RETURNING id
                    """
                ),
                {
                    "employee_id": employee_id,
                    "shift_date": shift_date,
                    "turno": emp["turno"],
                    "entry_at": ts,
                    "entry_guard": guardia.strip(),
                    "entry_status": status,
                    "late_reason": motivo_retardo.strip(),
                    "vehicle_expected": 1 if emp["tiene_vehiculo"] else 0,
                    "vehicle_entered": 1 if vehicle_required else 0,
                    "vehicle_front_entry": front_path,
                    "vehicle_trunk_entry": trunk_path,
                    "incident": observaciones.strip(),
                    "created_at": ts,
                    "updated_at": ts,
                }
            )
            record_id = str(result.scalar_one())
            audit(conn, "attendance", record_id, "ENTRY", guardia.strip())
            return {"ok": True, "message": f"Entrada registrada: {status}", "record_id": record_id}

        if movimiento == "salida":
            open_att = fetch_one(
                conn,
                "SELECT * FROM attendance WHERE employee_id = :employee_id AND exit_at IS NULL ORDER BY id DESC LIMIT 1",
                {"employee_id": employee_id}
            )
            if not open_att:
                return JSONResponse(status_code=400, content={"ok": False, "message": "No hay entrada abierta para registrar salida"})

            status = exit_status(open_att["turno"], open_att["shift_date"], dt)
            if status == "Salida temprana" and not motivo_salida_temprana.strip():
                return JSONResponse(status_code=400, content={"ok": False, "message": "Salida temprana: captura motivo"})

            conn.execute(
                text(
                    """
                    UPDATE attendance
                    SET exit_at=:exit_at,
                        exit_guard=:exit_guard,
                        exit_status=:exit_status,
                        early_reason=:early_reason,
                        vehicle_front_exit=:vehicle_front_exit,
                        vehicle_trunk_exit=:vehicle_trunk_exit,
                        updated_at=:updated_at,
                        incident=:incident
                    WHERE id=:id
                    """
                ),
                {
                    "exit_at": ts,
                    "exit_guard": guardia.strip(),
                    "exit_status": status,
                    "early_reason": motivo_salida_temprana.strip(),
                    "vehicle_front_exit": front_path,
                    "vehicle_trunk_exit": trunk_path,
                    "updated_at": ts,
                    "incident": observaciones.strip() or open_att["incident"],
                    "id": open_att["id"],
                }
            )
            audit(conn, "attendance", str(open_att["id"]), "EXIT", guardia.strip())
            return {"ok": True, "message": f"Salida registrada: {status}", "record_id": open_att["id"]}

    return JSONResponse(status_code=400, content={"ok": False, "message": "Movimiento inválido"})


@app.get("/healthz")
def healthz():
    return {
        "ok": True,
        "app": APP_NAME,
        "db": engine.dialect.name,
        "storage": str(DATA_DIR),
    }
