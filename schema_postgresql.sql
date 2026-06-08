-- Esquema PostgreSQL para Control QR Asistencias
-- La app crea estas tablas automáticamente al iniciar, pero este archivo sirve como referencia.

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
);

CREATE TABLE IF NOT EXISTS attendance (
    id SERIAL PRIMARY KEY,
    employee_id TEXT NOT NULL REFERENCES employees(id),
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
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS users (
    username TEXT PRIMARY KEY,
    password_hash TEXT NOT NULL,
    role TEXT DEFAULT 'Admin',
    active INTEGER DEFAULT 1,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS user_sessions (
    token_hash TEXT PRIMARY KEY,
    username TEXT NOT NULL REFERENCES users(username),
    expires_at TEXT NOT NULL,
    created_at TEXT NOT NULL
);

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
);

CREATE TABLE IF NOT EXISTS audit_log (
    id SERIAL PRIMARY KEY,
    table_name TEXT NOT NULL,
    record_id TEXT NOT NULL,
    action TEXT NOT NULL,
    user_name TEXT DEFAULT '',
    field_name TEXT DEFAULT '',
    old_value TEXT DEFAULT '',
    new_value TEXT DEFAULT '',
    reason TEXT DEFAULT '',
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_attendance_employee_open ON attendance(employee_id, exit_at);
CREATE INDEX IF NOT EXISTS idx_attendance_shift_date ON attendance(shift_date);
CREATE INDEX IF NOT EXISTS idx_audit_record ON audit_log(table_name, record_id);

CREATE INDEX IF NOT EXISTS idx_sessions_username ON user_sessions(username);
CREATE INDEX IF NOT EXISTS idx_sessions_expires ON user_sessions(expires_at);
