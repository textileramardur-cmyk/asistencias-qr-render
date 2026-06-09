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

function setMovement(movement, reload = false) {
  currentMovement = movement;
  const input = byId('movementType');
  if (input) input.value = movement;
  const fm = byId('formMovement');
  if (fm) fm.value = movement;
  const autoText = byId('autoMovementText');
  if (autoText) autoText.innerText = movement === 'entrada' ? 'Entrada' : 'Salida';
  const autoCard = byId('autoMovementCard');
  if (autoCard) {
    autoCard.classList.remove('salida');
    if (movement === 'salida') autoCard.classList.add('salida');
  }
  const chip = byId('movementChip');
  if (chip) chip.innerText = movement === 'entrada' ? 'Entrada' : 'Salida';
  const submit = byId('submitBtn');
  if (submit) submit.innerText = movement === 'entrada' ? 'Registrar entrada' : 'Registrar salida';
  if (reload && currentEmployee) loadEmployee();
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

async function parseJsonResponse(res) {
  const contentType = res.headers.get('content-type') || '';
  if (!contentType.includes('application/json')) {
    return { ok: false, message: res.status === 401 ? 'Sesión vencida. Inicia sesión nuevamente.' : 'Respuesta inesperada del servidor.' };
  }
  return await res.json();
}

function updateActiveGuardBox(activeGuard) {
  const box = byId('activeGuardBox');
  const name = byId('activeGuardName');
  const since = byId('activeGuardSince');
  if (!box || !name || !since || !activeGuard) return;
  box.classList.remove('missing');
  name.innerText = activeGuard.display || activeGuard.guard_display || 'Vigilante activo';
  since.innerText = activeGuard.started_at ? `Desde ${fmtDateTime(activeGuard.started_at)}` : 'Turno de vigilancia activo';
}

async function scanCode() {
  const code = byId('employeeId').value.trim();
  const resultBox = byId('resultBox');
  if (resultBox) hide(resultBox);
  if (!code) return;

  const res = await fetch(`/api/scan/${encodeURIComponent(code)}`);
  const data = await parseJsonResponse(res);
  if (!data.ok) {
    alert(data.message || 'Código no válido');
    return;
  }

  if (data.type === 'guard') {
    updateActiveGuardBox(data.active_guard);
    const status = byId('scannerStatus');
    if (status) status.innerText = data.message || 'Vigilante activo actualizado';
    byId('employeeId').value = '';
    currentEmployee = null;
    currentPreview = null;
    hide(byId('employeeCard'));
    return;
  }

  paintEmployeeData(data);
}

async function loadEmployee() {
  // Compatibilidad con botones viejos: ahora cualquier código pasa por el scanner inteligente.
  return scanCode();
}

function paintEmployeeData(data) {
  currentEmployee = data.employee;
  currentPreview = data.preview;
  if (data.active_guard) updateActiveGuardBox(data.active_guard);
  setMovement(data.next_movement || currentPreview?.movement || 'entrada', false);

  show(byId('employeeCard'));
  byId('empName').innerText = currentEmployee.nombre;
  byId('empMeta').innerText = `${currentEmployee.area || 'Sin área'} · ${currentEmployee.puesto || 'Sin puesto'}`;
  byId('empId').innerText = currentEmployee.id;
  byId('empTurno').innerText = currentEmployee.turno;
  byId('empVehicle').innerText = currentEmployee.tiene_vehiculo ? 'Sí' : 'No';
  byId('empOpen').innerText = data.has_open_attendance ? 'Sí' : 'No';
  byId('formEmployeeId').value = currentEmployee.id;

  const status = byId('empStatus');
  status.innerText = currentEmployee.estado;
  status.className = 'badge ' + (currentEmployee.estado === 'Activo' ? 'ok' : 'bad');

  const vehicleCheck = byId('vehicleCheck');
  if (vehicleCheck) vehicleCheck.checked = Boolean(currentEmployee.tiene_vehiculo);

  paintEvaluation(currentPreview);
  byId('employeeCard').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}


async function submitAttendance(event) {
  event.preventDefault();
  const form = event.target;
  byId('formMovement').value = currentMovement;

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
  const data = await parseJsonResponse(res);

  resultBox.className = 'result ' + (data.ok ? 'ok' : 'error');
  resultBox.innerText = data.message || (data.ok ? 'Registro guardado' : 'No se pudo registrar');

  if (data.ok) {
    form.reset();
    byId('employeeId').value = '';
    currentEmployee = null;
    currentPreview = null;
    hide(byId('employeeCard'));
    setMovement('entrada', false);
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
        scanCode();
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
