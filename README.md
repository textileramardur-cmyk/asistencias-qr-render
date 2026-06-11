{% extends "base.html" %}
{% block content %}
<section class="page-head row-head">
  <div>
    <h2>Recalcular reglas de asistencia</h2>
    <p>Actualiza registros ya capturados/importados usando la configuración actual de turnos. Útil cuando cambias tolerancias después de una carga.</p>
  </div>
  <div class="actions">
    <a class="btn secondary" href="/turnos">Ver turnos</a>
    <a class="btn secondary" href="/retardos?fecha={{ fecha_inicio }}">Ver retardos</a>
  </div>
</section>

<section class="card">
  <h3>Rango a recalcular</h3>
  <form method="get" action="/recalcular" class="grid-form compact-form">
    <label>Fecha inicio
      <input type="date" name="fecha_inicio" value="{{ fecha_inicio }}" required>
    </label>
    <label>Fecha fin
      <input type="date" name="fecha_fin" value="{{ fecha_fin }}" required>
    </label>
    <label>Turno
      <select name="turno">
        <option value="">Todos</option>
        {% for t in turnos %}
          <option value="{{ t.name }}" {% if turno == t.name %}selected{% endif %}>{{ t.name }}</option>
        {% endfor %}
      </select>
    </label>
    <div class="form-actions align-end">
      <button class="btn secondary" type="submit">Previsualizar rango</button>
    </div>
  </form>
</section>

<div class="kpi-grid compact">
  <div class="kpi"><small>Registros en rango</small><strong>{{ resumen.registros }}</strong></div>
  <div class="kpi"><small>Retardos actuales</small><strong>{{ resumen.retardos }}</strong></div>
  <div class="kpi"><small>Salidas tempranas</small><strong>{{ resumen.tempranas }}</strong></div>
  <div class="kpi"><small>Extras</small><strong>{{ resumen.extras }}</strong></div>
</div>

<section class="card warning-card">
  <h3>Aplicar recálculo</h3>
  <p>Esto recalcula <strong>retardo, minutos de retardo, salida temprana, extra y límites de horario</strong> con las reglas actuales del turno. No cambia horas de entrada/salida.</p>
  <p class="muted">Ejemplo: si antes importaste con tolerancia de 10 minutos y ahora configuraste 5 minutos, este proceso ajusta los registros del rango para que usen 5 minutos. Sí, el sistema por fin deja de fingir que el pasado no cambió de regla.</p>
  <form method="post" action="/recalcular" class="grid-form compact-form">
    <input type="hidden" name="fecha_inicio" value="{{ fecha_inicio }}">
    <input type="hidden" name="fecha_fin" value="{{ fecha_fin }}">
    <input type="hidden" name="turno" value="{{ turno }}">
    <label>Confirmación
      <input name="confirmar" placeholder="Escribe RECALCULAR" autocomplete="off" required>
    </label>
    <div class="form-actions align-end">
      <button class="btn danger" type="submit">Recalcular rango</button>
    </div>
  </form>
</section>

{% if result %}
<section class="card">
  <h3>Resultado del recálculo</h3>
  <div class="kpi-grid compact">
    <div class="kpi"><small>Registros revisados</small><strong>{{ result.scanned }}</strong></div>
    <div class="kpi"><small>Registros modificados</small><strong>{{ result.changed_records }}</strong></div>
    <div class="kpi"><small>Campos ajustados</small><strong>{{ result.changed_fields }}</strong></div>
    <div class="kpi"><small>Errores</small><strong>{{ result.total_errors }}</strong></div>
  </div>
  <div class="table-wrap compact-table">
    <table>
      <thead><tr><th>Movimiento</th><th>Altas</th><th>Eliminados</th></tr></thead>
      <tbody>
        <tr><td>Retardos</td><td>{{ result.late_added }}</td><td>{{ result.late_removed }}</td></tr>
        <tr><td>Salidas tempranas</td><td>{{ result.early_added }}</td><td>{{ result.early_removed }}</td></tr>
        <tr><td>Extras</td><td>{{ result.extra_added }}</td><td>{{ result.extra_removed }}</td></tr>
      </tbody>
    </table>
  </div>
  {% if result.errors %}
    <h4>Errores detectados</h4>
    <ul class="clean-list error-list">
      {% for e in result.errors %}<li>{{ e }}</li>{% endfor %}
    </ul>
  {% endif %}
  <div class="actions">
    <a class="btn primary" href="/retardos?fecha={{ fecha_inicio }}">Ver retardos</a>
    <a class="btn secondary" href="/exportar/asistencias.xlsx?fecha_inicio={{ fecha_inicio }}&fecha_fin={{ fecha_fin }}">Exportar asistencias</a>
  </div>
</section>
{% endif %}
{% endblock %}
