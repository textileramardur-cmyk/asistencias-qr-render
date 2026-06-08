-- Referencia de esquema para PostgreSQL. La app crea/migra tablas automáticamente al iniciar.

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

CREATE TABLE IF NOT EXISTS shift_settings (
  name TEXT PRIMARY KEY,
  entry_time TEXT NOT NULL,
  exit_time TEXT NOT NULL,
  crosses_midnight INTEGER DEFAULT 0,
  work_days TEXT DEFAULT '1,2,3,4,5',
  entry_tolerance_minutes INTEGER DEFAULT 10,
  exit_tolerance_minutes INTEGER DEFAULT 10,
  extra_after_minutes INTEGER DEFAULT 30,
  auto_close_enabled INTEGER DEFAULT 1,
  provisional_close_time TEXT DEFAULT '02:00',
  max_open_minutes INTEGER DEFAULT 1080,
  active INTEGER DEFAULT 1,
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
  incident TEXT DEFAULT '',
  scheduled_entry_at TEXT DEFAULT '',
  scheduled_exit_at TEXT DEFAULT '',
  entry_limit_at TEXT DEFAULT '',
  exit_early_limit_at TEXT DEFAULT '',
  extra_limit_at TEXT DEFAULT '',
  entry_tolerance_minutes INTEGER DEFAULT 10,
  exit_tolerance_minutes INTEGER DEFAULT 10,
  extra_after_minutes INTEGER DEFAULT 30,
  late_minutes INTEGER DEFAULT 0,
  early_minutes INTEGER DEFAULT 0,
  extra_minutes INTEGER DEFAULT 0,
  late_justified INTEGER DEFAULT 0,
  early_justified INTEGER DEFAULT 0,
  extra_authorized INTEGER DEFAULT 0,
  provisional_exit INTEGER DEFAULT 0,
  review_required INTEGER DEFAULT 0,
  review_status TEXT DEFAULT '',
  auto_closed_at TEXT DEFAULT '',
  anulled INTEGER DEFAULT 0,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS users (
  username TEXT PRIMARY KEY,
  password_hash TEXT NOT NULL,
  role TEXT DEFAULT 'Admin',
  active INTEGER DEFAULT 1,
  display_name TEXT DEFAULT '',
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

CREATE TABLE IF NOT EXISTS system_events (
  id SERIAL PRIMARY KEY,
  level TEXT DEFAULT 'info',
  module TEXT DEFAULT '',
  event_type TEXT DEFAULT '',
  user_name TEXT DEFAULT '',
  message TEXT DEFAULT '',
  detail TEXT DEFAULT '',
  created_at TEXT NOT NULL
);
