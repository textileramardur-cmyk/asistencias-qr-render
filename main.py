{% extends "base.html" %}
{% block content %}
<section class="page-head row-head">
  <div>
    <h2>Vigilantes</h2>
    <p>Códigos QR para cambio de turno de vigilancia. El último vigilante escaneado queda como responsable activo.</p>
  </div>
  <div class="actions">
    <a class="btn secondary" href="/vigilancia">Ir a vigilancia</a>
  </div>
</section>

<section class="card soft-card">
  <h3>Vigilante activo</h3>
  {% if active_guard %}
    <p><strong>{{ active_guard.display }}</strong></p>
    <p class="muted">Desde {{ short_datetime(active_guard.started_at) }} · Cambiado por {{ active_guard.changed_by }}</p>
  {% else %}
    <p class="muted">No hay vigilante activo. Escanea un QR de vigilante en /vigilancia.</p>
  {% endif %}
</section>

<section class="card">
  <h3>Alta / edición rápida</h3>
  <form method="post" action="/vigilantes/guardar" class="grid-form">
    <label>Código
      <input name="code" placeholder="VIG-ALTIMA-1" required />
    </label>
    <label>Alias
      <input name="alias" placeholder="Altima 1" required />
    </label>
    <label>Nombre
      <input name="nombre" placeholder="David" />
    </label>
    <label>Estado
      <select name="active"><option value="1">Activo</option><option value="0">Inactivo</option></select>
    </label>
    <label>QR
      <select name="qr_activo"><option value="1">Activo</option><option value="0">Inactivo</option></select>
    </label>
    <label>Observaciones
      <input name="observaciones" placeholder="Opcional" />
    </label>
    <div class="form-actions"><button class="btn primary" type="submit">Guardar vigilante</button></div>
  </form>
</section>

<section class="card">
  <div class="table-wrap">
    <table>
      <thead>
        <tr><th>Código</th><th>Alias</th><th>Nombre</th><th>Estado</th><th>QR</th><th>Acciones</th></tr>
      </thead>
      <tbody>
      {% for g in guards %}
        <tr>
          <td><strong>{{ g.code }}</strong></td>
          <td>{{ g.alias }}</td>
          <td>{{ g.nombre }}</td>
          <td>{{ 'Activo' if g.active else 'Inactivo' }}</td>
          <td>{{ 'Activo' if g.qr_activo else 'Inactivo' }}</td>
          <td class="table-actions wide-actions">
            <a href="/vigilantes/{{ g.code }}.png" target="_blank">QR PNG</a>
            <form method="post" action="/vigilantes/{{ g.code }}/estado" class="inline-form">
              <input type="hidden" name="active" value="{{ 0 if g.active else 1 }}" />
              <button class="link-btn" type="submit">{{ 'Inactivar' if g.active else 'Activar' }}</button>
            </form>
            {% set user = get_current_user(request) %}
            {% if user and user.role == 'Supremo' %}
            <form method="post" action="/vigilantes/{{ g.code }}/borrar" class="inline-form" onsubmit="return confirm('Borrado definitivo de vigilante. ¿Seguro?')">
              <input name="confirmacion" value="BORRAR" type="hidden" />
              <button class="link-btn danger-link" type="submit">Borrar</button>
            </form>
            {% endif %}
          </td>
        </tr>
      {% else %}
        <tr><td colspan="6">No hay vigilantes registrados.</td></tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
</section>
{% endblock %}
