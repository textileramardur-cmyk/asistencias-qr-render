let currentEmployee = null;
let currentPreview = null;
let scanner = null;
let scannerRunning = false;
let currentMovement = 'entrada';

function byId(id) { return document.getElementById(id); }
function show(el) { if (el) el.classList.remove('hidden'); }
function hide(el) { if (el) el.classList.add('hidden'); }
function pad(n) { return String(n).padStart(2, '0'); }
function fmtDateTime(value) {
  if (!value) return '-';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value).slice(0, 16);
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function setMovement(movement) {
  currentMovement = movement;
  const input = byId('movementType');
  if (input) input.value = movement;
  const fm = byId('formMovement');
  if (fm) fm.value = movement;
  document.querySelectorAll('.quick-card').forEach(btn => btn.classList.remove('active'));
  const buttons = document.querySelectorAll('.quick-card');
  if (movement === 'entrada' && buttons[0]) buttons[0].classList.add('active');
  if (movement === 'salida' && buttons[1]) buttons[1].classList.add('active');
  const chip = byId('movementChip');
  if (chip) chip.innerText = movement === 'entrada' ? 'Entrada' : 'Salida';
  const submit = byId('submitBtn');
  if (submit) submit.innerText = movement === 'entrada' ? 'Registrar entrada' : 'Registrar salida';
  if (currentEmployee) loadEmployee();
}

function toggleManualPanel() {
  const panel = byId('manualPanel');
  if (!panel) return;
  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function syncPanelsByPreview() {
  const late = byId('latePanel');
  const early = byId('earlyPanel');
  hide(late); hide(early);
  if (!currentPreview) return;
  if (currentPreview.status === 'Retardo') show(late);
  if (currentPreview.status === 'Salida temprana') show(early);
}

function paintEvaluation(preview) {
  const box = byId('statusBox');
  const statusEl = byId('evalStatus');
  const detailEl = byId('evalDetail');
  if (!box || !preview) return;
  box.className = 'status-box';
  let detail = '';
  if (preview.status === 'Retardo') {
    box.classList.add('warn');
    detail = `Límite ${fmtDateTime(preview.entry_limit_at)} · ${preview.late_minutes || 0} min tarde · motivo obligatorio.`;
  } else if (preview.status === 'Salida temprana') {
    box.classList.add('bad');
    detail = `Límite temprano ${fmtDateTime(preview.exit_early_limit_at)} · ${preview.early_minutes || 0} min · motivo obligatorio.`;
  } else if (preview.status === 'Extra') {
    box.classList.add('extra');
    detail = `Salida programada ${fmtDateTime(preview.scheduled_exit_at)} · extra desde ${fmtDateTime(preview.extra_limit_at)} · ${preview.extra_minutes || 0} min.`;
  } else {
    detail = `Horario OK · entrada límite ${fmtDateTime(preview.entry_limit_at)} · extra después de ${fmtDateTime(preview.extra_limit_at)}.`;
  }
  statusEl.innerText = preview.status || '-';
  detailEl.innerText = detail;
  syncPanelsByPreview();
}

async function loadEmployee() {
  const employeeId = byId('employeeId').value.trim();
  const resultBox = byId('resultBox');
  if (resultBox) hide(resultBox);
  if (!employeeId) return;

  const res = await fetch(`/api/empleado/${encodeURIComponent(employeeId)}?movimiento=${encodeURIComponent(currentMovement)}`);
  const data = await res.json();
  if (!data.ok) {
    alert(data.message || 'Empleado no encontrado');
    return;
  }

  currentEmployee = data.employee;
  currentPreview = data.preview;

  show(byId('employeeCard'));
  byId('empName').innerText = currentEmployee.nombre;
  byId('empMeta').innerText = `${currentEmployee.area || 'Sin área'} · ${currentEmployee.puesto || 'Sin puesto'}`;
  byId('empId').innerText = currentEmployee.id;
  byId('empTurno').innerText = currentEmployee.turno;
  byId('empVehicle').innerText = currentEmployee.tiene_vehiculo ? 'Sí' : 'No';
  byId('empOpen').innerText = data.has_open_attendance ? 'Sí' : 'No';
  byId('formEmployeeId').value = currentEmployee.id;
  byId('formGuardia').value = byId('guardName').value || 'Vigilancia';

  const status = byId('empStatus');
  status.innerText = currentEmployee.estado;
  status.className = 'badge ' + (currentEmployee.estado === 'Activo' ? 'ok' : 'bad');

  const vehicleCheck = byId('vehicleCheck');
  if (vehicleCheck) vehicleCheck.checked = Boolean(currentEmployee.tiene_vehiculo);

  // Si el empleado tiene entrada abierta, sugerimos salida; si no, entrada. Menos clics, menos oportunidad para el caos.
  if (data.has_open_attendance && currentMovement !== 'salida') setMovement('salida');
  else if (!data.has_open_attendance && currentMovement !== 'entrada') setMovement('entrada');
  else paintEvaluation(currentPreview);

  byId('employeeCard').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

async function submitAttendance(event) {
  event.preventDefault();
  const form = event.target;
  byId('formMovement').value = currentMovement;
  byId('formGuardia').value = byId('guardName').value || 'Vigilancia';

  if (currentPreview?.status === 'Retardo' && !byId('lateReason').value) {
    alert('Hay retardo. El motivo es obligatorio. Triste, pero razonable.');
    return;
  }
  if (currentPreview?.status === 'Salida temprana' && !byId('earlyReason').value) {
    alert('Hay salida temprana. El motivo es obligatorio.');
    return;
  }

  const resultBox = byId('resultBox');
  resultBox.className = 'result hidden';

  const formData = new FormData(form);
  if (!byId('vehicleCheck')?.checked) formData.set('vehiculo', '0');

  const res = await fetch('/api/registro', { method: 'POST', body: formData });
  const data = await res.json();

  resultBox.className = 'result ' + (data.ok ? 'ok' : 'error');
  resultBox.innerText = data.message || (data.ok ? 'Registro guardado' : 'No se pudo registrar');

  if (data.ok) {
    form.reset();
    byId('guardName').value = 'Vigilancia';
    byId('employeeId').value = '';
    currentEmployee = null;
    currentPreview = null;
    hide(byId('employeeCard'));
    setMovement('entrada');
    byId('scannerStatus').innerText = data.message;
  }
}

async function toggleScanner() {
  const readerId = 'reader';
  const status = byId('scannerStatus');
  if (scannerRunning && scanner) {
    await scanner.stop();
    scannerRunning = false;
    if (status) status.innerText = 'Cámara detenida. Puedes capturar ID manual.';
    return;
  }

  if (!window.Html5Qrcode) {
    alert('No cargó el lector QR. Usa captura manual del ID. La web siendo la web, una historia vieja.');
    return;
  }

  scanner = new Html5Qrcode(readerId);
  try {
    await scanner.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 280, height: 280 } },
      (decodedText) => {
        byId('employeeId').value = decodedText.trim();
        if (status) status.innerText = `QR leído: ${decodedText.trim()}`;
        loadEmployee();
        scanner.stop();
        scannerRunning = false;
      }
    );
    scannerRunning = true;
    if (status) status.innerText = 'Cámara activa. Centra el QR en el marco.';
  } catch (err) {
    if (status) status.innerText = 'No se pudo abrir cámara. Revisa permisos o usa ID manual.';
    alert('No se pudo abrir la cámara. Revisa permisos del navegador. También puedes escribir el ID manualmente.');
  }
}
