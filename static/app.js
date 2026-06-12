let currentEmployee = null;
let currentPreview = null;
let scanner = null;
let scannerRunning = false;
let currentMovement = 'entrada';
let qrBusy = false;
let noQrMode = false;
let noQrReasonText = '';
let feedbackTimer = null;

function byId(id) { return document.getElementById(id); }
function show(el) { if (el) el.classList.remove('hidden'); }
function hide(el) { if (el) el.classList.add('hidden'); }
function pad(n) { return String(n).padStart(2, '0'); }

function fmtMinutesDuration(value) {
  const n = Math.max(0, Math.round(Number(value || 0)));
  const h = Math.floor(n / 60);
  const m = n % 60;
  if (h === 0) return `00:${String(m).padStart(2, '0')}`;
  return `${h}:${String(m).padStart(2, '0')}`;
}

function fmtDateTime(value) {
  if (!value) return '-';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value).slice(0, 16);
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function todayTime() {
  const d = new Date();
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function beep(kind = 'ok') {
  try {
    if (navigator.vibrate) navigator.vibrate(kind === 'error' ? [120, 70, 120] : 80);
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    if (!AudioCtx) return;
    const ctx = new AudioCtx();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = 'sine';
    osc.frequency.value = kind === 'error' ? 180 : (kind === 'warn' ? 440 : 760);
    gain.gain.value = 0.04;
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    setTimeout(() => { osc.stop(); ctx.close(); }, kind === 'error' ? 280 : 130);
  } catch (_) {}
}

function showFeedback(kind, title, text) {
  const overlay = byId('bigFeedback');
  const card = overlay?.querySelector('.guard-feedback-card');
  const icon = byId('feedbackIcon');
  const titleEl = byId('feedbackTitle');
  const textEl = byId('feedbackText');
  if (!overlay || !card) return;
  card.className = `guard-feedback-card ${kind}`;
  if (icon) icon.innerText = kind === 'error' ? '!' : (kind === 'warn' ? '!' : '✓');
  if (titleEl) titleEl.innerText = title || 'Listo';
  if (textEl) textEl.innerText = text || '';
  show(overlay);
  beep(kind === 'error' ? 'error' : (kind === 'warn' ? 'warn' : 'ok'));
  clearTimeout(feedbackTimer);
  feedbackTimer = setTimeout(() => hide(overlay), kind === 'error' ? 3600 : 2200);
}

function setMovement(movement, reload = false) {
  currentMovement = movement || 'entrada';
  const input = byId('movementType');
  if (input) input.value = currentMovement;
  const fm = byId('formMovement');
  if (fm) fm.value = currentMovement;

  const chip = byId('movementChip');
  if (chip) {
    chip.innerText = currentMovement === 'entrada' ? 'Entrada' : 'Salida';
    chip.className = 'badge ' + (currentMovement === 'entrada' ? 'ok' : 'neutral');
  }
  const title = byId('movementTitle');
  if (title) title.innerText = currentMovement === 'entrada' ? 'ENTRADA' : 'SALIDA';
  const subtitle = byId('movementSubtitle');
  if (subtitle) subtitle.innerText = 'El sistema lo decidió automáticamente.';
  const banner = byId('movementBanner');
  if (banner) {
    banner.className = 'movement-banner ' + (currentMovement === 'salida' ? 'salida' : 'entrada');
  }
  const submit = byId('submitBtn');
  if (submit) submit.innerText = currentMovement === 'entrada' ? 'CONFIRMAR ENTRADA' : 'CONFIRMAR SALIDA';
  if (reload && currentEmployee) loadEmployee();
}

function setGuardInstruction(active = true, display = '') {
  const zone = byId('guardCameraZone');
  const pill = byId('stepGuardPill');
  const title = byId('stepGuardTitle');
  const text = byId('stepGuardText');
  const main = byId('guardMainInstruction');
  if (zone) zone.classList.toggle('needs-guard', !active);
  if (pill) pill.className = `step-pill ${active ? 'ok' : 'warn'}`;
  if (title) title.innerText = active ? 'Vigilante activo' : 'Falta vigilante';
  if (text) text.innerText = active ? 'Ya puedes escanear empleados.' : 'Escanea primero el QR del vigilante. Si no, el sistema no registra empleados.';
  if (main) {
    const p = main.querySelector('p');
    const h1 = main.querySelector('h1');
    if (p) p.innerText = active ? 'LISTO PARA CAPTURAR' : 'PASO 1';
    if (h1) h1.innerText = active ? 'Escanea QR de empleado' : 'Escanea QR de vigilante';
  }
}

function toggleManualPanel() {
  const panel = byId('manualPanel');
  if (!panel) return;
  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  const input = byId('employeeId');
  if (input) setTimeout(() => input.focus(), 300);
}

function toggleNoQrPanel() {
  const panel = byId('noQrPanel');
  if (!panel) return;
  panel.classList.toggle('hidden');
  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  if (!panel.classList.contains('hidden')) {
    const q = byId('noQrSearch');
    if (q) setTimeout(() => q.focus(), 300);
  }
}

function setQuickReason(targetId, value) {
  const field = byId(targetId);
  if (field) {
    field.value = value;
    field.dispatchEvent(new Event('change'));
  }
}

function syncVehicleChoice() {
  const checked = document.querySelector('input[name="vehiculo_choice"]:checked');
  const hidden = byId('vehicleHidden');
  if (hidden) hidden.value = checked?.value || '0';
}

function setVehicleChoice(value) {
  const wanted = String(value ? 1 : 0);
  const radio = document.querySelector(`input[name="vehiculo_choice"][value="${wanted}"]`);
  if (radio) radio.checked = true;
  syncVehicleChoice();
}

function syncPanelsByPreview() {
  const late = byId('latePanel');
  const early = byId('earlyPanel');
  const lunch = byId('lunchPanel');
  hide(late); hide(early); hide(lunch);
  if (!currentPreview) return;
  if (currentPreview.status === 'Retardo') show(late);
  if (currentPreview.status === 'Salida temprana') { show(early); show(lunch); }
}

function paintEvaluation(preview) {
  const box = byId('statusBox');
  const statusEl = byId('evalStatus');
  const detailEl = byId('evalDetail');
  if (!box || !preview) return;
  box.className = 'status-box guard-eval-box';
  let detail = '';
  if (preview.requires_confirmation) box.classList.add('warn');

  if (preview.status === 'Retardo') {
    box.classList.add('warn');
    detail = `Límite ${fmtDateTime(preview.entry_limit_at)} · ${fmtMinutesDuration(preview.late_minutes || 0)} tarde · motivo obligatorio.`;
  } else if (preview.status === 'Salida temprana') {
    box.classList.add('bad');
    detail = `Límite temprano ${fmtDateTime(preview.exit_early_limit_at)} · ${fmtMinutesDuration(preview.early_minutes || 0)} · motivo obligatorio.`;
  } else if (preview.status === 'Extra') {
    box.classList.add('extra');
    detail = `Salida programada ${fmtDateTime(preview.scheduled_exit_at)} · extra desde ${fmtDateTime(preview.extra_limit_at)} · ${fmtMinutesDuration(preview.extra_minutes || 0)}.`;
  } else if (preview.status === 'Reentrada') {
    box.classList.add('warn');
    detail = preview.message || 'Ya tuvo movimiento en este turno. Se pedirá confirmación.';
  } else {
    box.classList.add('ok');
    detail = `Horario OK · entrada límite ${fmtDateTime(preview.entry_limit_at)} · extra después de ${fmtDateTime(preview.extra_limit_at)}.`;
  }
  statusEl.innerText = preview.requires_confirmation ? `${preview.status || '-'} · requiere confirmación` : (preview.status || '-');
  detailEl.innerText = preview.requires_confirmation ? `${preview.message || 'Movimiento sensible: confirma antes de guardar.'} ${detail}` : detail;
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
  setGuardInstruction(true, name.innerText);
}

async function scanCode() {
  const input = byId('employeeId');
  const code = input?.value.trim() || '';
  const resultBox = byId('resultBox');
  if (resultBox) hide(resultBox);
  if (!code || qrBusy) return;
  qrBusy = true;
  const status = byId('scannerStatus');
  if (status) status.innerText = 'Validando código...';

  try {
    const res = await fetch(`/api/scan/${encodeURIComponent(code)}`);
    const data = await parseJsonResponse(res);
    if (!data.ok) {
      showFeedback('error', 'NO SE REGISTRÓ', data.message || 'Código no válido');
      if (status) status.innerText = data.message || 'Código no válido';
      return;
    }

    if (data.type === 'guard') {
      updateActiveGuardBox(data.active_guard);
      if (status) status.innerText = data.message || 'Vigilante activo actualizado';
      if (input) input.value = '';
      currentEmployee = null;
      currentPreview = null;
      hide(byId('employeeCard'));
      showFeedback('ok', 'VIGILANTE ACTIVO', data.active_guard?.display || data.message || 'Listo');
      return;
    }

    paintEmployeeData(data);
    if (status) status.innerText = `${data.employee?.nombre || 'Empleado'} detectado. Confirma el movimiento.`;
    beep('ok');
  } catch (err) {
    showFeedback('error', 'ERROR', 'No se pudo validar el código. Revisa conexión.');
  } finally {
    setTimeout(() => { qrBusy = false; }, 900);
  }
}

async function loadEmployee() {
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

  setVehicleChoice(Boolean(currentEmployee.tiene_vehiculo));
  paintEvaluation(currentPreview);
  byId('employeeCard').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function cancelEmployee() {
  const form = byId('attendanceForm');
  if (form) form.reset();
  byId('employeeId').value = '';
  currentEmployee = null;
  currentPreview = null;
  noQrMode = false;
  noQrReasonText = '';
  hide(byId('employeeCard'));
  setMovement('entrada', false);
  const status = byId('scannerStatus');
  if (status) status.innerText = 'Cancelado. Escanea el siguiente QR.';
}

async function submitAttendance(event) {
  event.preventDefault();
  const form = event.target;
  byId('formMovement').value = currentMovement;

  if (!currentEmployee) {
    showFeedback('error', 'FALTA EMPLEADO', 'Escanea o busca un empleado antes de confirmar.');
    return;
  }
  if (currentPreview?.status === 'Retardo' && !byId('lateReason').value) {
    showFeedback('warn', 'FALTA MOTIVO', 'Hay retardo. Selecciona un motivo.');
    return;
  }
  if (currentPreview?.status === 'Salida temprana' && !byId('earlyReason').value) {
    showFeedback('warn', 'FALTA MOTIVO', 'Hay salida temprana. Selecciona un motivo.');
    return;
  }
  if (currentPreview?.status === 'Salida temprana' && !byId('lunchTaken').value) {
    showFeedback('warn', 'FALTA COMIDA', 'Indica si tomó su hora de comida.');
    return;
  }

  const confirmInput = byId('formConfirmacionOperativa');
  if (confirmInput) confirmInput.value = '';
  if (currentPreview?.requires_confirmation) {
    const ok = window.confirm((currentPreview.message || 'Este movimiento requiere doble confirmación.') + '\n\nConfirma solo si el trabajador realmente está entrando o reingresando.');
    if (!ok) return;
    if (confirmInput) confirmInput.value = 'CONFIRMAR';
  }

  const obs = byId('observaciones');
  if (noQrMode) {
    const reason = noQrReasonText || byId('noQrReason')?.value || '';
    if (!reason) {
      showFeedback('warn', 'FALTA MOTIVO', 'Para Sin QR el motivo es obligatorio.');
      return;
    }
    const prefix = `SIN QR: ${reason}`;
    obs.value = obs.value && !obs.value.startsWith('SIN QR:') ? `${prefix}. ${obs.value}` : prefix;
  }

  const resultBox = byId('resultBox');
  resultBox.className = 'result hidden';
  const formData = new FormData(form);
  syncVehicleChoice();
  formData.set('vehiculo', byId('vehicleHidden')?.value || '0');

  const submit = byId('submitBtn');
  if (submit) { submit.disabled = true; submit.innerText = 'GUARDANDO...'; }

  try {
    const res = await fetch('/api/registro', { method: 'POST', body: formData });
    const data = await parseJsonResponse(res);

    resultBox.className = 'result ' + (data.ok ? 'ok' : 'error');
    resultBox.innerText = data.message || (data.ok ? 'Registro guardado' : 'No se pudo registrar');

    if (data.ok) {
      const movementWord = currentMovement === 'entrada' ? 'ENTRADA REGISTRADA' : 'SALIDA REGISTRADA';
      const kind = data.status === 'Retardo' || data.status === 'Salida temprana' || data.status === 'Reentrada' ? 'warn' : 'ok';
      showFeedback(kind, movementWord, `${currentEmployee.nombre} · ${todayTime()} · ${data.message || data.status || 'OK'}`);
      form.reset();
      byId('employeeId').value = '';
      currentEmployee = null;
      currentPreview = null;
      noQrMode = false;
      noQrReasonText = '';
      hide(byId('employeeCard'));
      setMovement('entrada', false);
      const status = byId('scannerStatus');
      if (status) status.innerText = 'Registro guardado. Escanea el siguiente empleado.';
      loadLastMovements();
    } else {
      showFeedback('error', 'NO SE REGISTRÓ', data.message || 'No se pudo registrar');
    }
  } catch (err) {
    showFeedback('error', 'ERROR', 'No se pudo guardar. Revisa conexión.');
  } finally {
    if (submit) { submit.disabled = false; setMovement(currentMovement, false); }
  }
}

async function toggleScanner() {
  const readerId = 'reader';
  const status = byId('scannerStatus');
  const btnText = byId('cameraButtonText');
  const cameraZone = byId('guardCameraZone');

  const setStatus = (msg) => { if (status) status.innerText = msg; };
  const markCameraOff = () => {
    scannerRunning = false;
    cameraZone?.classList.remove('camera-active');
    if (btnText) btnText.innerText = 'Activar cámara';
  };

  async function stopScannerInstance() {
    if (!scanner) return;
    try {
      if (scannerRunning) await scanner.stop();
    } catch (_) {}
    try { await scanner.clear(); } catch (_) {}
    scanner = null;
    markCameraOff();
  }

  if (scannerRunning && scanner) {
    await stopScannerInstance();
    setStatus('Cámara detenida. Puedes capturar ID manual.');
    return;
  }

  if (!window.Html5Qrcode) {
    showFeedback('error', 'CÁMARA NO DISPONIBLE', 'El lector QR no cargó. Usa captura manual del ID.');
    return;
  }

  const host = window.location.hostname || '';
  const isLocal = ['localhost', '127.0.0.1', '0.0.0.0'].includes(host);
  if (window.location.protocol !== 'https:' && !isLocal) {
    setStatus('La cámara requiere HTTPS. Abre la app desde el enlace https de Render.');
    showFeedback('error', 'CÁMARA BLOQUEADA', 'El navegador solo permite cámara en sitios HTTPS.');
    return;
  }

  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    setStatus('Este navegador no permite cámara aquí. Usa ID manual.');
    showFeedback('error', 'SIN SOPORTE DE CÁMARA', 'Prueba Safari/Chrome actualizado o captura ID manual.');
    return;
  }

  const onDecoded = async (decodedText) => {
    if (qrBusy) return;
    byId('employeeId').value = decodedText.trim();
    setStatus('QR leído. Validando...');
    await scanCode();
    await stopScannerInstance();
  };

  const baseOptions = [
    { fps: 10, qrbox: { width: 320, height: 320 }, disableFlip: true },
    { fps: 8, qrbox: { width: 280, height: 280 }, disableFlip: true },
    { fps: 6, disableFlip: true }
  ];

  async function startWith(cameraConfig, label) {
    for (const options of baseOptions) {
      await stopScannerInstance();
      scanner = new Html5Qrcode(readerId, { verbose: false });
      await scanner.start(cameraConfig, options, onDecoded);
      scannerRunning = true;
      cameraZone?.classList.add('camera-active');
      if (btnText) btnText.innerText = 'Detener cámara';
      setStatus(`Cámara activa${label ? ' · ' + label : ''}. Acerca el QR al centro.`);
      return true;
    }
    return false;
  }

  try {
    setStatus('Solicitando permiso de cámara...');

    // Primero intentamos con lista de cámaras. Es más compatible que exigir resolución HD.
    let cameras = [];
    try {
      cameras = await Html5Qrcode.getCameras();
    } catch (err) {
      // Algunos navegadores no entregan lista antes de permiso. Seguimos con constraints simples.
      cameras = [];
    }

    if (cameras && cameras.length) {
      const backWords = ['back', 'rear', 'environment', 'trasera', 'posterior', 'atrás', 'atras'];
      const backCamera = cameras.find(c => backWords.some(w => (c.label || '').toLowerCase().includes(w)));
      const selected = backCamera || cameras[cameras.length - 1] || cameras[0];
      await startWith(selected.id, selected.label || 'cámara disponible');
      return;
    }

    // Fallback muy compatible: sin width/height obligatorios, porque eso rompe permisos en varios celulares.
    const constraintFallbacks = [
      { facingMode: { ideal: 'environment' } },
      { facingMode: 'environment' },
      { facingMode: { ideal: 'user' } },
      true
    ];

    let lastError = null;
    for (const cfg of constraintFallbacks) {
      try {
        await startWith(cfg, 'modo compatible');
        return;
      } catch (err) {
        lastError = err;
      }
    }
    throw lastError;
  } catch (err) {
    console.error('Camera start error:', err);
    await stopScannerInstance();
    const name = err?.name || '';
    let message = 'No se pudo abrir cámara. Revisa permisos o usa ID manual.';
    if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
      message = 'Permiso de cámara denegado. Actívalo en el candado del navegador.';
    } else if (name === 'NotFoundError' || name === 'DevicesNotFoundError') {
      message = 'No se encontró cámara disponible en este dispositivo.';
    } else if (name === 'NotReadableError' || name === 'TrackStartError') {
      message = 'La cámara está ocupada por otra app o el navegador la bloqueó.';
    } else if (name === 'OverconstrainedError' || name === 'ConstraintNotSatisfiedError') {
      message = 'La cámara no aceptó la configuración. Intenta de nuevo en modo compatible.';
    }
    setStatus(message);
    showFeedback('error', 'SIN CÁMARA', message);
  }
}

async function loadLastMovements() {
  const box = byId('lastMovementsList');
  if (!box) return;
  try {
    const res = await fetch('/api/vigilancia/ultimos');
    const data = await parseJsonResponse(res);
    if (!data.ok || !data.rows?.length) {
      box.innerHTML = '<div class="last-empty">Aún no hay movimientos recientes.</div>';
      return;
    }
    box.innerHTML = data.rows.map(row => {
      const cls = row.status === 'Retardo' || row.status === 'Salida temprana' ? 'warn' : (row.movement === 'Salida' ? 'exit' : 'entry');
      return `<div class="last-movement-item ${cls}">
        <div><strong>${row.time || '-'}</strong><small>${row.movement || '-'}</small></div>
        <span>${row.nombre || row.employee_id}</span>
        <em>${row.status || 'OK'}</em>
      </div>`;
    }).join('');
  } catch (_) {
    box.innerHTML = '<div class="last-empty">No se pudieron cargar los últimos movimientos.</div>';
  }
}

async function searchNoQrEmployees() {
  const q = byId('noQrSearch')?.value.trim() || '';
  const reason = byId('noQrReason')?.value || '';
  const out = byId('noQrResults');
  if (!out) return;
  if (!reason) {
    showFeedback('warn', 'FALTA MOTIVO', 'Selecciona primero el motivo de Sin QR.');
    return;
  }
  if (q.length < 2) {
    out.innerHTML = '<div class="last-empty">Escribe mínimo 2 caracteres.</div>';
    return;
  }
  out.innerHTML = '<div class="last-empty">Buscando...</div>';
  try {
    const res = await fetch(`/api/vigilancia/buscar?q=${encodeURIComponent(q)}`);
    const data = await parseJsonResponse(res);
    if (!data.ok || !data.rows?.length) {
      out.innerHTML = '<div class="last-empty">Sin resultados.</div>';
      return;
    }
    out.innerHTML = data.rows.map(emp => `<button class="noqr-result" type="button" onclick="selectNoQrEmployee('${String(emp.id).replace(/'/g, '')}', '${String(emp.nombre || '').replace(/'/g, '')}')">
      <strong>${emp.nombre}</strong><span>ID ${emp.id} · ${emp.area || 'Sin área'} · ${emp.turno || '-'}</span>
    </button>`).join('');
  } catch (_) {
    out.innerHTML = '<div class="last-empty">Error buscando empleados.</div>';
  }
}

async function selectNoQrEmployee(id, name) {
  const reason = byId('noQrReason')?.value || '';
  if (!reason) {
    showFeedback('warn', 'FALTA MOTIVO', 'Selecciona el motivo de Sin QR.');
    return;
  }
  noQrMode = true;
  noQrReasonText = reason;
  byId('employeeId').value = id;
  const obs = byId('observaciones');
  if (obs) obs.value = `SIN QR: ${reason}`;
  await scanCode();
  const status = byId('scannerStatus');
  if (status) status.innerText = `Sin QR: ${name || id}. Confirma el movimiento.`;
}
