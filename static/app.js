let currentEmployee = null;
let scanner = null;
let scannerRunning = false;

function byId(id) { return document.getElementById(id); }
function show(el) { el.classList.remove('hidden'); }
function hide(el) { el.classList.add('hidden'); }

function syncMovementUI() {
  const movement = byId('movementType')?.value || 'entrada';
  const formMovement = byId('formMovement');
  if (formMovement) formMovement.value = movement;
  const late = byId('latePanel');
  const early = byId('earlyPanel');
  if (late && early) {
    if (movement === 'entrada') { show(late); hide(early); }
    else { hide(late); show(early); }
  }
}

function syncVehicleUI() {
  const checked = byId('vehicleCheck')?.checked;
  const vehiclePhotos = byId('vehiclePhotos');
  if (!vehiclePhotos) return;
  checked ? show(vehiclePhotos) : hide(vehiclePhotos);
}

async function loadEmployee() {
  const employeeId = byId('employeeId').value.trim();
  const resultBox = byId('resultBox');
  if (resultBox) hide(resultBox);
  if (!employeeId) return;

  const res = await fetch(`/api/empleado/${encodeURIComponent(employeeId)}`);
  const data = await res.json();
  if (!data.ok) {
    alert(data.message || 'Empleado no encontrado');
    return;
  }

  currentEmployee = data.employee;
  byId('employeeCard').classList.remove('hidden');
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

  const photo = byId('employeePhoto');
  if (photo) photo.innerHTML = '👤';

  const vehicleCheck = byId('vehicleCheck');
  vehicleCheck.checked = Boolean(currentEmployee.tiene_vehiculo);
  syncVehicleUI();
  syncMovementUI();
}

async function submitAttendance(event) {
  event.preventDefault();
  const form = event.target;
  byId('formMovement').value = byId('movementType').value;
  byId('formGuardia').value = byId('guardName').value || 'Vigilancia';

  const resultBox = byId('resultBox');
  resultBox.className = 'result hidden';

  const formData = new FormData(form);
  if (!byId('vehicleCheck').checked) {
    formData.set('vehiculo', '0');
  }

  const res = await fetch('/api/registro', { method: 'POST', body: formData });
  const data = await res.json();

  resultBox.className = 'result ' + (data.ok ? 'ok' : 'error');
  resultBox.innerText = data.message || (data.ok ? 'Registro guardado' : 'No se pudo registrar');

  if (data.ok) {
    form.reset();
    byId('movementType').value = 'entrada';
    byId('guardName').value = 'Vigilancia';
    byId('employeeId').value = currentEmployee?.id || '';
    hide(byId('employeeCard'));
    setTimeout(() => window.location.href = '/monitor', 900);
  }
}

async function toggleScanner() {
  const readerId = 'reader';
  if (scannerRunning && scanner) {
    await scanner.stop();
    scannerRunning = false;
    return;
  }

  if (!window.Html5Qrcode) {
    alert('No cargó el lector QR. Usa captura manual del ID. Así es la web: promete magia y luego pide permisos.');
    return;
  }

  scanner = new Html5Qrcode(readerId);
  try {
    await scanner.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 240, height: 240 } },
      (decodedText) => {
        byId('employeeId').value = decodedText.trim();
        loadEmployee();
        scanner.stop();
        scannerRunning = false;
      }
    );
    scannerRunning = true;
  } catch (err) {
    alert('No se pudo abrir la cámara. Revisa permisos del navegador. También puedes escribir el ID manualmente.');
  }
}
