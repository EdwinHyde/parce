// ════════════════════════════════════════════════════════
//  CONFIGURACIÓN
// ════════════════════════════════════════════════════════
const API_KEY           = "AIzaSyA1GsdvskVoIQDSOg6EI8aEQW5MnQsFeoE";
const SPREADSHEET_ID    = "1ytP_bXfbjRsA05KJYM8HAxQvTvnAeHgz09r4QhMy-Rs";
const SHEET_ESTUDIANTES = "asignaturas_estudiantes";
const SHEET_PERIODOS    = "Periodos";

const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbygBydUcKQWczkcs4jL3BshWNY1npLtsqK9IbNA72nIgpRhmjHo3tKnCVry0AcwvGIf/exec";

const PESO_SABER = 0.35;
const PESO_HACER = 0.35;
const PESO_SER   = 0.30;

// ════════════════════════════════════════════════════════
//  ESTADO GLOBAL
// ════════════════════════════════════════════════════════
let nombreDocente       = "";
let sedeDocente         = "";
let todosLosEstudiantes = [];
let headersEstudiantes  = [];
let estudiantesActuales = [];
let asignaturaActiva    = "";
let gradoActivo         = "";
let periodoActivo       = "";
let notasPrevias        = {};
let periodosBloqueados  = {};

// Matriz de inputs para navegación 2D [fila][col]
// Cada fila tiene: [s1, s2, s3, s4, h1, h2, h3, h4, se1, se2, rec]
let matrizInputs = [];

// ════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════
function getUrlParameter(name) {
  name = name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]');
  const regex = new RegExp('[\\?&]' + name + '=([^&#]*)');
  const results = regex.exec(location.search);
  return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
}

function parseFecha(valor) {
  if (!valor && valor !== 0) return null;
  if (typeof valor === 'number' || (String(valor).match(/^\d+$/) && String(valor).length < 6)) {
    const base = new Date(1899, 11, 30);
    base.setDate(base.getDate() + Number(valor));
    base.setHours(0,0,0,0);
    return base;
  }
  const str = String(valor).trim();
  const dmy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) return new Date(+dmy[3], +dmy[2]-1, +dmy[1]);
  const ymd = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) return new Date(+ymd[1], +ymd[2]-1, +ymd[3]);
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function fetchSheet(sheetName) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(sheetName)}?key=${API_KEY}`;
  return fetch(url).then(r => r.json());
}

function promedio(vals) {
  const nums = vals.map(v => parseFloat(v)).filter(v => !isNaN(v) && v >= 1 && v <= 5);
  if (nums.length === 0) return null;
  return Math.round((nums.reduce((a, b) => a + b, 0) / nums.length) * 10) / 10;
}

function colorNota(val) {
  if (val === null) return '';
  const n = parseFloat(val);
  if (n >= 3.5) return 'nota-alta';
  if (n >= 3.0) return 'nota-media';
  return 'nota-baja';
}

function setSaveStatus(msg, tipo) {
  const el = document.getElementById('save-status');
  el.textContent = msg;
  el.className = 'save-status' + (tipo ? ' ' + tipo : '');
}

function mostrarEstado(id) {
  ['state-loading', 'state-empty', 'table-section'].forEach(s => {
    document.getElementById(s).style.display = 'none';
  });
  if (id) {
    document.getElementById(id).style.display = id === 'table-section' ? 'block' : 'flex';
  }
}

const METADATOS = ['no_identificador','nombres_apellidos','grado','sede','estado'];

const ORDEN_GRADOS = [
  'Preescolar','Primero','Segundo','Tercero','Cuarto','Quinto',
  'Sexto','Séptimo','Octavo','Noveno','Décimo','Undécimo'
];

function ordenarGrados(lista) {
  return [...lista].sort((a, b) => {
    const ia = ORDEN_GRADOS.indexOf(a);
    const ib = ORDEN_GRADOS.indexOf(b);
    if (ia === -1 && ib === -1) return a.localeCompare(b);
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });
}

function colsAsignatura() {
  return headersEstudiantes
    .map((h, i) => ({ nombre: h, idx: i }))
    .filter(c => !METADATOS.includes(c.nombre.toLowerCase().trim()));
}

// ════════════════════════════════════════════════════════
//  ENVÍO A APPS SCRIPT
// ════════════════════════════════════════════════════════
function enviarAAppsScript(registros) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify({ registros });
    fetch(APPS_SCRIPT_URL, {
      method:   'POST',
      headers:  { 'Content-Type': 'text/plain' },
      body:     payload,
      redirect: 'follow'
    })
    .then(r => r.text().then(text => {
      try { resolve(JSON.parse(text)); }
      catch { resolve({ status: 'ok', guardados: registros.length, mensaje: registros.length + ' registro(s) enviados.' }); }
    }))
    .catch(() => enviarConIframe(payload).then(resolve).catch(reject));
  });
}

function enviarConIframe(payload) {
  return new Promise((resolve) => {
    const iframeName = 'sianario_frame_' + Date.now();
    const iframe = document.createElement('iframe');
    iframe.name  = iframeName;
    iframe.style.display = 'none';
    document.body.appendChild(iframe);

    const form = document.createElement('form');
    form.method  = 'POST';
    form.action  = APPS_SCRIPT_URL;
    form.target  = iframeName;
    form.enctype = 'application/x-www-form-urlencoded';

    const input = document.createElement('input');
    input.type  = 'hidden';
    input.name  = 'payload';
    input.value = payload;
    form.appendChild(input);
    document.body.appendChild(form);

    const limpiar = () => {
      try { document.body.removeChild(iframe); } catch(e) {}
      try { document.body.removeChild(form);   } catch(e) {}
    };

    const timeout = setTimeout(() => {
      limpiar();
      resolve({ status: 'ok', mensaje: 'Datos enviados. Confirma en tu Sheets.' });
    }, 12000);

    iframe.onload = () => {
      clearTimeout(timeout);
      try { resolve(JSON.parse(iframe.contentDocument?.body?.innerText || '')); }
      catch { resolve({ status: 'ok', mensaje: 'Datos enviados. Confirma en tu Sheets.' }); }
      limpiar();
    };
    form.submit();
  });
}

// ════════════════════════════════════════════════════════
//  CARGAR NOTAS PREVIAS
// ════════════════════════════════════════════════════════
async function cargarNotasPrevias(periodo, asignatura, grado) {
  try {
    const url = `${APPS_SCRIPT_URL}?action=getCalificaciones` +
                `&periodo=${encodeURIComponent(periodo)}` +
                `&asignatura=${encodeURIComponent(asignatura)}` +
                `&grado=${encodeURIComponent(grado)}`;
    const res  = await fetch(url);
    const data = await res.json();
    return data.status === 'ok' ? (data.registros || {}) : {};
  } catch { return {}; }
}

// ════════════════════════════════════════════════════════
//  INICIO
// ════════════════════════════════════════════════════════
nombreDocente = getUrlParameter('nombre');
sedeDocente   = getUrlParameter('sede');
document.getElementById('header-nombre').textContent = nombreDocente || 'Docente';
document.getElementById('header-sede').textContent   = sedeDocente   || '';

document.getElementById('btn-back').addEventListener('click', () => {
  const params = `?nombre=${encodeURIComponent(nombreDocente)}&sede=${encodeURIComponent(sedeDocente)}`;
  window.location.href = 'docente_inicio.html' + params;
});

document.getElementById('btn-logout').addEventListener('click', () => {
  if (confirm('¿Estás seguro de que deseas cerrar sesión?')) {
    window.location.href = 'index.html';
  }
});

// ════════════════════════════════════════════════════════
//  INICIALIZACIÓN
// ════════════════════════════════════════════════════════
async function inicializar() {
  document.getElementById('loader-grado').classList.add('visible');
  try {
    const [data, dataPer] = await Promise.all([
      fetchSheet(SHEET_ESTUDIANTES),
      fetchSheet(SHEET_PERIODOS)
    ]);

    // Procesar periodos — bloqueo por CIERRE SISTEMA
    if (!dataPer.error && dataPer.values && dataPer.values.length > 1) {
      const hPer    = dataPer.values[0].map(h => (h||'').trim());
      const iNum    = hPer.findIndex(h => h.toUpperCase().includes('PERIODO'));
      const iInicio = hPer.findIndex(h => h.toUpperCase().includes('INICIO'));
      const iCierre = hPer.findIndex(h =>
        h.toUpperCase().includes('CIERRE') && h.toUpperCase().includes('SISTEMA')
      );
      const hoy = new Date(); hoy.setHours(0,0,0,0);
      dataPer.values.slice(1).forEach(row => {
        const num    = String(row[iNum] || '').trim();
        const inicio = iInicio >= 0 ? parseFecha(row[iInicio]) : null;
        const cierre = iCierre >= 0 ? parseFecha(row[iCierre]) : null;
        if (num) {
          const antesDeApertura = inicio ? hoy < inicio : false;
          const despuesDeCierre = cierre ? hoy > cierre : false;
          periodosBloqueados[num] = antesDeApertura || despuesDeCierre;
        }
      });
    }

    // Procesar estudiantes
    if (!data.error && data.values && data.values.length > 1) {
      headersEstudiantes  = data.values[0].map(h => (h||'').trim());
      todosLosEstudiantes = data.values.slice(1);
    }
  } catch(err) {
    console.error('Error inicializando:', err);
  } finally {
    document.getElementById('loader-grado').classList.remove('visible');
    const selPer = document.getElementById('sel-periodo');
    if (selPer.value) selPer.dispatchEvent(new Event('change'));
  }
}
inicializar();

// ════════════════════════════════════════════════════════
//  PASO 1 — Periodo → Grados
// ════════════════════════════════════════════════════════
document.getElementById('sel-periodo').addEventListener('change', function () {
  const periodo  = this.value;
  const selGrado = document.getElementById('sel-grado');
  const selAsig  = document.getElementById('sel-asignatura');

  selGrado.innerHTML = '<option value="">— Elige el grado —</option>';
  selGrado.disabled  = true;
  selAsig.innerHTML  = '<option value="">— Primero elige el grado —</option>';
  selAsig.disabled   = true;
  mostrarEstado(null);
  document.getElementById('selection-summary').style.display = 'none';

  if (!periodo) return;
  if (!todosLosEstudiantes.length) {
    document.getElementById('loader-grado').classList.add('visible');
    return;
  }

  const iGrado  = headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO');
  const iEstado = headersEstudiantes.findIndex(h => h.toUpperCase() === 'ESTADO');
  const cols    = colsAsignatura();

  const filaDocente = todosLosEstudiantes.filter(row => {
    const estado = (row[iEstado] || '').trim().toLowerCase();
    if (estado === 'inactivo' || estado === 'retirado') return false;
    return cols.some(c => (row[c.idx] || '').trim() === nombreDocente.trim());
  });

  const gradosUnicos = ordenarGrados(
    [...new Set(filaDocente.map(r => (r[iGrado] || '').trim()))].filter(Boolean)
  );

  if (!gradosUnicos.length) {
    selGrado.innerHTML = '<option value="">Sin grados asignados</option>';
    return;
  }

  selGrado.innerHTML =
    '<option value="">— Elige el grado —</option>' +
    '<option value="__todos__">— Todos los grados —</option>' +
    gradosUnicos.map(g => `<option value="${g}">${g}</option>`).join('');
  selGrado.disabled = false;
});

// ════════════════════════════════════════════════════════
//  PASO 2 — Grado → Asignaturas
// ════════════════════════════════════════════════════════
document.getElementById('sel-grado').addEventListener('change', function () {
  const grado   = this.value;
  const selAsig = document.getElementById('sel-asignatura');

  selAsig.innerHTML = '<option value="">— Elige la asignatura —</option>';
  selAsig.disabled  = true;
  mostrarEstado(null);

  if (!grado) return;

  const iEstado = headersEstudiantes.findIndex(h => h.toUpperCase() === 'ESTADO');
  const iGrado  = headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO');
  const cols    = colsAsignatura();

  // Si es "todos los grados", buscar en todos los estudiantes activos del docente
  const filasFiltradas = todosLosEstudiantes.filter(row => {
    const estadoFila = (row[iEstado] || '').trim().toLowerCase();
    if (estadoFila === 'inactivo' || estadoFila === 'retirado') return false;
    const gradoOk = grado === '__todos__'
      ? true
      : (row[iGrado] || '').trim().toLowerCase() === grado.toLowerCase();
    return gradoOk;
  });

  const asignaturasDocente = cols
    .filter(c => filasFiltradas.some(row => (row[c.idx] || '').trim() === nombreDocente.trim()))
    .map(c => c.nombre);

  if (!asignaturasDocente.length) {
    selAsig.innerHTML = '<option value="">Sin asignaturas para este grado</option>';
    return;
  }

  selAsig.innerHTML = '<option value="">— Elige la asignatura —</option>' +
    asignaturasDocente.map(a => `<option value="${a}">${a}</option>`).join('');
  selAsig.disabled = false;
});

// ════════════════════════════════════════════════════════
//  PASO 3 — Asignatura → cargar tabla
// ════════════════════════════════════════════════════════
document.getElementById('sel-asignatura').addEventListener('change', function () {
  if (this.value) cargarEstudiantes();
});

// ════════════════════════════════════════════════════════
//  CARGAR ESTUDIANTES
// ════════════════════════════════════════════════════════
async function cargarEstudiantes() {
  periodoActivo    = document.getElementById('sel-periodo').value;
  gradoActivo      = document.getElementById('sel-grado').value;
  asignaturaActiva = document.getElementById('sel-asignatura').value;
  if (!periodoActivo || !gradoActivo || !asignaturaActiva) return;

  mostrarEstado('state-loading');
  document.getElementById('state-loading').querySelector('p').textContent = 'Cargando estudiantes...';

  const iGrado      = headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO');
  const iEstado     = headersEstudiantes.findIndex(h => h.toUpperCase() === 'ESTADO');
  const iAsigActiva = headersEstudiantes.findIndex(h => h === asignaturaActiva);

  estudiantesActuales = todosLosEstudiantes.filter(row => {
    const estadoFila = (row[iEstado]     || '').trim().toLowerCase();
    if (estadoFila === 'inactivo' || estadoFila === 'retirado') return false;
    const gradoOk = gradoActivo === '__todos__'
      ? true
      : (row[iGrado] || '').trim().toLowerCase() === gradoActivo.toLowerCase();
    return gradoOk && (row[iAsigActiva] || '').trim() === nombreDocente.trim();
  });

  // Ordenar por grado académico cuando son todos los grados
  if (gradoActivo === '__todos__') {
    const iGradoSort = headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO');
    const iNombreSort = headersEstudiantes.findIndex(h =>
      h.trim().toLowerCase() === 'nombres y apellidos' ||
      h.trim().toLowerCase() === 'nombres_apellidos'
    );
    estudiantesActuales.sort((a, b) => {
      const ga = ordenarGrados([(a[iGradoSort]||'').trim(), (b[iGradoSort]||'').trim()]);
      if (ga[0] !== (a[iGradoSort]||'').trim()) return 1;
      if (ga[0] !== (b[iGradoSort]||'').trim()) return -1;
      return ((a[iNombreSort]||'').localeCompare(b[iNombreSort]||''));
    });
  }

  if (!estudiantesActuales.length) {
    mostrarEstado(null);
    document.getElementById('state-empty-msg').textContent =
      `No se encontraron estudiantes activos de ${gradoActivo} en ${asignaturaActiva}.`;
    mostrarEstado('state-empty');
    return;
  }

  document.getElementById('state-loading').querySelector('p').textContent = 'Recuperando notas anteriores...';

  if (gradoActivo !== '__todos__') {
    notasPrevias = await cargarNotasPrevias(periodoActivo, asignaturaActiva, gradoActivo);
  } else {
    // Obtener los grados únicos presentes en los estudiantes actuales
    const iGradoPrev = headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO');
    const gradosPresentes = [...new Set(
      estudiantesActuales.map(r => (r[iGradoPrev] || '').trim()).filter(Boolean)
    )];

    // Cargar notas de cada grado en paralelo y combinar
    const resultados = await Promise.all(
      gradosPresentes.map(g => cargarNotasPrevias(periodoActivo, asignaturaActiva, g))
    );
    notasPrevias = Object.assign({}, ...resultados);
  }

  mostrarEstado(null);
  renderTabla();

  document.getElementById('chip-periodo').textContent    = `Periodo ${periodoActivo}`;
  document.getElementById('chip-grado').textContent =
    gradoActivo === '__todos__' ? 'Todos los grados' : gradoActivo;
  document.getElementById('chip-asignatura').textContent = asignaturaActiva;
  document.getElementById('selection-summary').style.display = 'flex';
}

document.getElementById('btn-cargar').addEventListener('click', cargarEstudiantes);

// ════════════════════════════════════════════════════════
//  RENDERIZAR TABLA
// ════════════════════════════════════════════════════════
function renderTabla() {
  const iNombre = headersEstudiantes.findIndex(h =>
    h.trim().toLowerCase() === 'nombres y apellidos' ||
    h.trim().toLowerCase() === 'nombres_apellidos'   ||
    h.trim().toLowerCase() === 'nombres y apeliidos'
  );
  const iSede = headersEstudiantes.findIndex(h => h.toUpperCase() === 'SEDE');
  const iId   = headersEstudiantes.findIndex(h =>
    h.toUpperCase() === 'ID_ESTUDIANTE' ||
    h.toUpperCase() === 'NO DOCUMENTO'  ||
    h.toUpperCase() === 'NO_DOCUMENTO'
  );

  // Mostrar/ocultar columna Grado según selección
  const mostrarGrado = gradoActivo === '__todos__';
  const thGrado = document.getElementById('th-grado-col');
  if (mostrarGrado && !thGrado) {
    const th = document.createElement('th');
    th.id        = 'th-grado-col';
    th.className = 'th-fixed th-sede';
    th.textContent = 'Grado';
    // Insertar después de th-sede (3ª posición, índice 2)
    const thead = document.querySelector('.calificaciones-table thead tr');
    thead.insertBefore(th, thead.children[3]);
  } else if (!mostrarGrado && thGrado) {
    thGrado.remove();
  }

  const tbody = document.getElementById('tabla-body');
  tbody.innerHTML = '';
  matrizInputs    = [];

  estudiantesActuales.forEach((row, idx) => {
    const nombre  = (iNombre >= 0 ? row[iNombre] : '') || 'Sin nombre';
    const sede    = (row[iSede]   || '—').trim();
    const idEstud = (row[iId]     || '').trim();
    const prev    = notasPrevias[idEstud] || {};

    function getPrev(clave) {
      const claveNorm = clave.toUpperCase();
      const entrada   = Object.entries(prev).find(([k]) => k.toUpperCase() === claveNorm);
      if (!entrada) return '';
      const val = parseFloat(entrada[1]);
      return (!isNaN(val) && val >= 1 && val <= 5) ? entrada[1] : '';
    }

    const vs1  = getPrev('SABER_1');
    const vs2  = getPrev('SABER_2');
    const vs3  = getPrev('SABER_3');
    const vs4  = getPrev('SABER_4');
    const vh1  = getPrev('HACER_1');
    const vh2  = getPrev('HACER_2');
    const vh3  = getPrev('HACER_3');
    const vh4  = getPrev('HACER_4');
    const vse1 = getPrev('SER_1');
    const vse2 = getPrev('SER_2');
    const vrec = getPrev('RECUPERACION');
    const tienePrevias = Object.keys(prev).length > 0;

    const tr = document.createElement('tr');
    tr.dataset.rowIdx  = idx;
    tr.dataset.idEstud = idEstud;

    tr.innerHTML = `
      <td class="td-num">${idx + 1}</td>
      <td class="td-nombre" title="${nombre.trim()}">${nombre.trim()}</td>
      <td class="td-sede">${sede}</td>
      ${mostrarGrado ? `<td class="td-sede" style="color:var(--iean-dark);font-weight:600;">${(row[headersEstudiantes.findIndex(h=>h.toUpperCase()==='GRADO')]||'').trim()}</td>` : ''}
      <td><input type="number" class="nota-input saber" data-comp="saber" data-sub="1" min="1" max="5" step="0.1" placeholder="—" value="${vs1}"></td>
      <td><input type="number" class="nota-input saber" data-comp="saber" data-sub="2" min="1" max="5" step="0.1" placeholder="—" value="${vs2}"></td>
      <td><input type="number" class="nota-input saber" data-comp="saber" data-sub="3" min="1" max="5" step="0.1" placeholder="—" value="${vs3}"></td>
      <td><input type="number" class="nota-input saber" data-comp="saber" data-sub="4" min="1" max="5" step="0.1" placeholder="—" value="${vs4}"></td>
      <td class="td-cons" id="cons-s-${idx}">—</td>
      <td><input type="number" class="nota-input hacer" data-comp="hacer" data-sub="1" min="1" max="5" step="0.1" placeholder="—" value="${vh1}"></td>
      <td><input type="number" class="nota-input hacer" data-comp="hacer" data-sub="2" min="1" max="5" step="0.1" placeholder="—" value="${vh2}"></td>
      <td><input type="number" class="nota-input hacer" data-comp="hacer" data-sub="3" min="1" max="5" step="0.1" placeholder="—" value="${vh3}"></td>
      <td><input type="number" class="nota-input hacer" data-comp="hacer" data-sub="4" min="1" max="5" step="0.1" placeholder="—" value="${vh4}"></td>
      <td class="td-cons" id="cons-h-${idx}">—</td>
      <td><input type="number" class="nota-input ser" data-comp="ser" data-sub="1" min="1" max="5" step="0.1" placeholder="—" value="${vse1}"></td>
      <td><input type="number" class="nota-input ser" data-comp="ser" data-sub="2" min="1" max="5" step="0.1" placeholder="—" value="${vse2}"></td>
      <td class="td-cons" id="cons-sr-${idx}">—</td>
      <td class="rec-cell" id="rec-cell-${idx}">
        <input type="number" class="nota-input rec" data-comp="rec" data-sub="1" min="1" max="5" step="0.1" placeholder="—" value="${vrec}">
      </td>
      <td class="td-prom" id="prom-${idx}">—</td>
    `;

    if (tienePrevias) {
      tr.querySelector('.td-nombre').style.borderLeft = '3px solid var(--iean-main)';
    }

    tbody.appendChild(tr);
    calcularFila(idx);

    // Recoger inputs de esta fila en orden y añadir a la matriz
    const inputsFila = [...tr.querySelectorAll('.nota-input')];
    matrizInputs.push(inputsFila);

    inputsFila.forEach((input, colIdx) => {
      input.addEventListener('input', () => {
        validarInput(input);
        calcularFila(idx);
        setSaveStatus('');
      });

      // Navegación con teclado
      input.addEventListener('keydown', e => manejarTeclado(e, idx, colIdx));

      // Pegado desde Excel
      input.addEventListener('paste', e => {
        e.preventDefault();
        manejarPegado(e, input);
      });
    });
  });

  const tienePrevias = Object.keys(notasPrevias).length > 0;
  document.getElementById('table-title').textContent = asignaturaActiva;
  document.getElementById('table-count').textContent =
    `${estudiantesActuales.length} estudiante${estudiantesActuales.length !== 1 ? 's' : ''}` +
    (tienePrevias ? ' · Notas previas cargadas' : '');

  mostrarEstado('table-section');

  const periodoCerrado = periodosBloqueados[periodoActivo] === true;
  if (periodoCerrado) {
    bloquearTabla();
    setSaveStatus('🔒 El sistema aún no está habilitado para ingresar calificaciones en este periodo.', 'error');
  } else {
    setSaveStatus(tienePrevias ? '📋 Notas del periodo cargadas. Edita y guarda cuando estés listo.' : '');
    // Enfocar el primer input
    if (matrizInputs.length > 0 && matrizInputs[0].length > 0) {
      setTimeout(() => matrizInputs[0][0].focus(), 100);
    }
  }
}

// ════════════════════════════════════════════════════════
//  NAVEGACIÓN CON TECLADO — flechas, Enter, Tab
// ════════════════════════════════════════════════════════
function manejarTeclado(e, fila, col) {
  const totalFilas = matrizInputs.length;
  const totalCols  = matrizInputs[fila] ? matrizInputs[fila].length : 0;
  let nuevaFila = fila, nuevaCol = col;

  switch (e.key) {
    case 'ArrowRight':
      if (col < totalCols - 1) nuevaCol = col + 1;
      else if (fila < totalFilas - 1) { nuevaFila = fila + 1; nuevaCol = 0; }
      e.preventDefault(); break;

    case 'ArrowLeft':
      if (col > 0) nuevaCol = col - 1;
      else if (fila > 0) { nuevaFila = fila - 1; nuevaCol = (matrizInputs[fila-1]?.length || 1) - 1; }
      e.preventDefault(); break;

    case 'ArrowDown':
    case 'Enter':
      if (fila < totalFilas - 1) nuevaFila = fila + 1;
      e.preventDefault(); break;

    case 'ArrowUp':
      if (fila > 0) nuevaFila = fila - 1;
      e.preventDefault(); break;

    case 'Tab':
      if (!e.shiftKey && col === totalCols - 1 && fila < totalFilas - 1) {
        e.preventDefault(); nuevaFila = fila + 1; nuevaCol = 0;
      } else if (e.shiftKey && col === 0 && fila > 0) {
        e.preventDefault(); nuevaFila = fila - 1; nuevaCol = (matrizInputs[fila-1]?.length || 1) - 1;
      } else return;
      break;

    default: return;
  }

  const dest = matrizInputs[nuevaFila]?.[nuevaCol];
  if (dest) { dest.focus(); dest.select(); }
}

// ════════════════════════════════════════════════════════
//  PEGADO DESDE EXCEL — horizontal y vertical
// ════════════════════════════════════════════════════════
function manejarPegado(e, inputOrigen) {
  const texto = (e.clipboardData || window.clipboardData).getData('text');
  if (!texto) return;

  const filas = texto
    .replace(/\r\n/g, '\n').replace(/\r/g, '\n')
    .split('\n').filter(f => f.trim() !== '');

  const totalFilas    = filas.length;
  const totalColumnas = filas[0] ? filas[0].split('\t').length : 1;

  const todosInputs = [...document.querySelectorAll('.nota-input:not([disabled])')];
  const idxOrigen   = todosInputs.indexOf(inputOrigen);
  if (idxOrigen < 0) return;

  const inputsPorFila = matrizInputs[0]?.length || 1;
  const filaOrigen    = Math.floor(idxOrigen / inputsPorFila);
  const colOrigen     = idxOrigen % inputsPorFila;

  function normalizarValor(val) {
    const v = val.trim().replace(',', '.');
    const n = parseFloat(v);
    if (isNaN(n) || n < 1 || n > 5) return '';
    return String(Math.round(n * 10) / 10);
  }

  const inputsAfectados = [];

  if (totalColumnas === 1 && totalFilas > 1) {
    // Vertical
    filas.forEach((fila, i) => {
      const val  = normalizarValor(fila.split('\t')[0]);
      const dest = todosInputs[(filaOrigen + i) * inputsPorFila + colOrigen];
      if (dest) { dest.value = val; inputsAfectados.push(dest); }
    });
  } else if (totalFilas === 1 && totalColumnas > 1) {
    // Horizontal
    filas[0].split('\t').forEach((celda, j) => {
      const val     = normalizarValor(celda);
      const destIdx = filaOrigen * inputsPorFila + (colOrigen + j);
      const dest    = todosInputs[destIdx];
      if (dest && Math.floor(destIdx / inputsPorFila) === filaOrigen) {
        dest.value = val; inputsAfectados.push(dest);
      }
    });
  } else if (totalFilas > 1 && totalColumnas > 1) {
    // Bloque
    filas.forEach((fila, i) => {
      fila.split('\t').forEach((celda, j) => {
        const val  = normalizarValor(celda);
        const dest = todosInputs[(filaOrigen + i) * inputsPorFila + (colOrigen + j)];
        if (dest) { dest.value = val; inputsAfectados.push(dest); }
      });
    });
  } else {
    const val = normalizarValor(filas[0]?.split('\t')[0] || '');
    inputOrigen.value = val;
    inputsAfectados.push(inputOrigen);
  }

  inputsAfectados.forEach(inp => {
    validarInput(inp);
    const trPadre = inp.closest('tr');
    if (trPadre) calcularFila(parseInt(trPadre.dataset.rowIdx));
    inp.dispatchEvent(new Event('input', { bubbles: true }));
  });

  setSaveStatus(`📋 ${inputsAfectados.length} celda(s) pegadas. Revisa y guarda.`, '');
  if (inputsAfectados.length > 0) inputsAfectados[inputsAfectados.length - 1].focus();
}

// ════════════════════════════════════════════════════════
//  BLOQUEAR TABLA
// ════════════════════════════════════════════════════════
function bloquearTabla() {
  document.querySelectorAll('.nota-input').forEach(input => {
    input.disabled = true;
    input.style.background  = '#f5f5f5';
    input.style.color       = '#9e9e9e';
    input.style.cursor      = 'not-allowed';
    input.style.borderColor = '#e0e0e0';
  });
  document.getElementById('btn-guardar').style.display        = 'none';
  document.getElementById('btn-guardar-bottom').style.display = 'none';
  document.getElementById('btn-limpiar').style.display        = 'none';

  const headerBar = document.querySelector('.table-header-bar');
  if (headerBar && !headerBar.querySelector('.lock-banner')) {
    const banner = document.createElement('div');
    banner.className = 'lock-banner';
    banner.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
        <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
      </svg>
      Periodo cerrado — modo de lectura`;
    headerBar.appendChild(banner);
  }
}

// ════════════════════════════════════════════════════════
//  CALCULAR FILA
// ════════════════════════════════════════════════════════
function calcularFila(idx) {
  const tr = document.querySelector(`tr[data-row-idx="${idx}"]`);
  if (!tr) return;

  const getVals = comp =>
    [...tr.querySelectorAll(`.nota-input[data-comp="${comp}"]`)].map(i => i.value);

  // Promedio EXACTO sin redondeo (solo se redondea al mostrar y al guardar)
  function promedioExacto(vals) {
    const nums = vals.map(v => parseFloat(v)).filter(v => !isNaN(v) && v >= 1 && v <= 5);
    if (nums.length === 0) return null;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  }

  const consSaberExacto = promedioExacto(getVals('saber'));
  const consHacerExacto = promedioExacto(getVals('hacer'));
  const consSerExacto   = promedioExacto(getVals('ser'));

  // Mostrar consolidados redondeados a 1 decimal (solo visual)
  document.getElementById(`cons-s-${idx}`).textContent  = consSaberExacto !== null ? consSaberExacto.toFixed(1) : '—';
  document.getElementById(`cons-h-${idx}`).textContent  = consHacerExacto !== null ? consHacerExacto.toFixed(1) : '—';
  document.getElementById(`cons-sr-${idx}`).textContent = consSerExacto   !== null ? consSerExacto.toFixed(1)   : '—';

  let promFinal = null;
  if (consSaberExacto !== null && consHacerExacto !== null && consSerExacto !== null) {
    // Calcular con valores EXACTOS y redondear SOLO el resultado final
    const resultadoExacto = consSaberExacto * PESO_SABER + consHacerExacto * PESO_HACER + consSerExacto * PESO_SER;
    promFinal = Math.round(resultadoExacto * 10) / 10;
  }

  const promCell = document.getElementById(`prom-${idx}`);
  const recCell  = document.getElementById(`rec-cell-${idx}`);

  if (promFinal !== null) {
    promCell.textContent = promFinal.toFixed(1);
    promCell.className   = `td-prom ${colorNota(promFinal)}`;
    if (promFinal < 3.0) {
      recCell.classList.add('activa');
    } else {
      recCell.classList.remove('activa');
      recCell.querySelector('.nota-input.rec').value = '';
    }
  } else {
    promCell.textContent = '—';
    promCell.className   = 'td-prom';
    recCell.classList.remove('activa');
  }
}

// ════════════════════════════════════════════════════════
//  VALIDAR INPUT
// ════════════════════════════════════════════════════════
function validarInput(input) {
  if (input.value === '') { input.classList.remove('invalid'); return true; }
  const v  = parseFloat(input.value);
  const ok = !isNaN(v) && v >= 1.0 && v <= 5.0;
  input.classList.toggle('invalid', !ok);
  return ok;
}

// ════════════════════════════════════════════════════════
//  LIMPIAR
// ════════════════════════════════════════════════════════
document.getElementById('btn-limpiar').addEventListener('click', () => {
  if (!confirm('¿Deseas limpiar todos los campos? Los datos no guardados se perderán.')) return;
  document.querySelectorAll('.nota-input').forEach(i => { i.value = ''; i.classList.remove('invalid'); });
  document.querySelectorAll('[id^="cons-"]').forEach(el => el.textContent = '—');
  document.querySelectorAll('[id^="prom-"]').forEach(el => { el.textContent = '—'; el.className = 'td-prom'; });
  document.querySelectorAll('.rec-cell').forEach(c => c.classList.remove('activa'));
  document.querySelectorAll('.td-nombre').forEach(c => c.style.borderLeft = '');
  setSaveStatus('');
});

// ════════════════════════════════════════════════════════
//  RECOPILAR DATOS
//  - Excluye estudiantes sin ninguna nota ingresada
//  - Mapea cada campo individual (saber1..4, hacer1..4, ser1..2, rec)
//    con las claves exactas que espera guardarEnCalificaciones en Apps Script
// ════════════════════════════════════════════════════════
function recopilarDatos() {
  const iNombre = headersEstudiantes.findIndex(h =>
    h.trim().toLowerCase() === 'nombres y apellidos' ||
    h.trim().toLowerCase() === 'nombres_apellidos'   ||
    h.trim().toLowerCase() === 'nombres y apeliidos'
  );
  const iId  = headersEstudiantes.findIndex(h =>
    h.toUpperCase() === 'ID_ESTUDIANTE' ||
    h.toUpperCase() === 'NO DOCUMENTO'  ||
    h.toUpperCase() === 'NO_DOCUMENTO'
  );
  const iSeq = headersEstudiantes.findIndex(h => h.toUpperCase() === 'NO_IDENTIFICADOR');
  const iSede = headersEstudiantes.findIndex(h => h.toUpperCase() === 'SEDE');

  const registros = [];

  document.querySelectorAll('#tabla-body tr').forEach((tr, idx) => {
    const row = estudiantesActuales[idx];
    if (!row) return;

    const getComp = comp =>
      [...tr.querySelectorAll(`.nota-input[data-comp="${comp}"]`)].map(i => i.value.trim());

    const saberVals = getComp('saber');  // [s1, s2, s3, s4]
    const hacerVals = getComp('hacer');  // [h1, h2, h3, h4]
    const serVals   = getComp('ser');    // [se1, se2]
    const recVal    = tr.querySelector('.nota-input[data-comp="rec"]')?.value.trim() || '';

    // CORRECCIÓN: excluir filas donde absolutamente ningún campo tiene valor
    const tieneAlgo = [...saberVals, ...hacerVals, ...serVals, recVal].some(v => v !== '');
    if (!tieneAlgo) return;

    // Promedio EXACTO sin redondeo prematuro (misma lógica que calcularFila)
    function promedioExacto(vals) {
      const nums = vals.map(v => parseFloat(v)).filter(v => !isNaN(v) && v >= 1 && v <= 5);
      if (nums.length === 0) return null;
      return nums.reduce((a, b) => a + b, 0) / nums.length;
    }

    const consSaberExacto = promedioExacto(saberVals);
    const consHacerExacto = promedioExacto(hacerVals);
    const consSerExacto   = promedioExacto(serVals);

    let promFinal = null;
    if (consSaberExacto !== null && consHacerExacto !== null && consSerExacto !== null) {
      // Calcular con valores EXACTOS y redondear SOLO el resultado final
      const resultadoExacto = consSaberExacto * PESO_SABER + consHacerExacto * PESO_HACER + consSerExacto * PESO_SER;
      promFinal = Math.round(resultadoExacto * 10) / 10;
    }

    // Nota final: si hay recuperación, se toma la más alta entre promedio y recuperación
    const recNum   = recVal !== '' ? parseFloat(recVal) : null;
    const notaFinal = (recNum !== null && promFinal !== null)
      ? Math.max(promFinal, recNum)
      : (promFinal !== null ? promFinal : recNum);

    let desempeno = '';
    if (notaFinal !== null) {
      if (notaFinal >= 4.6) desempeno = 'SUPERIOR';
      else if (notaFinal >= 4.0) desempeno = 'ALTO';
      else if (notaFinal >= 3.0) desempeno = 'BASICO';
      else desempeno = 'BAJO';
    }

    registros.push({
      periodo:    periodoActivo,
      fecha:      new Date().toLocaleDateString('es-CO'),
      docente:    nombreDocente,
      grado: gradoActivo === '__todos__'
        ? (row[headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO')] || '').trim()
        : gradoActivo,
      sede:       (row[iSede]   || '').trim(),
      asignatura: asignaturaActiva,
      id:    (row[iId]  || '').trim(),
      idSeq: (row[iSeq] || '').trim(),
      nombre:     (iNombre >= 0 ? (row[iNombre] || '') : '').trim(),
      // Campos individuales — coinciden exactamente con el switch de guardarEnCalificaciones
      saber1:    saberVals[0] || '',
      saber2:    saberVals[1] || '',
      saber3:    saberVals[2] || '',
      saber4:    saberVals[3] || '',
      consSaber: consSaberExacto !== null ? consSaberExacto.toFixed(1) : '',
      hacer1:    hacerVals[0] || '',
      hacer2:    hacerVals[1] || '',
      hacer3:    hacerVals[2] || '',
      hacer4:    hacerVals[3] || '',
      consHacer: consHacerExacto !== null ? consHacerExacto.toFixed(1) : '',
      ser1:      serVals[0] || '',
      ser2:      serVals[1] || '',
      consSer:   consSerExacto !== null ? consSerExacto.toFixed(1) : '',
      rec:       recVal,
      promFinal:  promFinal  !== null ? promFinal.toFixed(1)  : '',
      notaFinal:  notaFinal  !== null ? notaFinal.toFixed(1)  : '',
      desempeno
    });
  });

  return registros;
}

// ════════════════════════════════════════════════════════
//  GUARDAR
// ════════════════════════════════════════════════════════
async function guardarCalificaciones() {
  if (periodosBloqueados[periodoActivo] === true) {
    setSaveStatus('🔒 No se puede guardar: el periodo está cerrado.', 'error');
    return;
  }

  const invalidos = document.querySelectorAll('.nota-input.invalid');
  if (invalidos.length > 0) {
    setSaveStatus('Corrige las notas en rojo antes de guardar.', 'error');
    invalidos[0].focus();
    return;
  }

  const conNotas = [...document.querySelectorAll('.nota-input')].some(i => i.value.trim() !== '');
  if (!conNotas) {
    setSaveStatus('Ingresa al menos una nota antes de guardar.', 'error');
    return;
  }

  const btnTop = document.getElementById('btn-guardar');
  const btnBot = document.getElementById('btn-guardar-bottom');
  btnTop.disabled = btnBot.disabled = true;
  setSaveStatus('Guardando calificaciones...', '');

  const registros = recopilarDatos();

  if (!registros.length) {
    setSaveStatus('No hay filas con datos para guardar.', 'error');
    btnTop.disabled = btnBot.disabled = false;
    return;
  }

  try {
    const result = await enviarAAppsScript(registros);
    if (result.status === 'ok') {
      setSaveStatus(`✓ ${result.guardados || registros.length} registro(s) guardados correctamente.`, 'success');
    } else if (result.status === 'parcial') {
      setSaveStatus(`⚠️ Guardado parcial: ${result.mensaje}`, 'error');
      console.warn('Errores:', result.errores);
    } else {
      setSaveStatus(`Error: ${result.mensaje}`, 'error');
    }
  } catch(err) {
    setSaveStatus('Error al enviar datos. Verifica tu conexión.', 'error');
    console.error('Error al guardar:', err);
  }

  btnTop.disabled = btnBot.disabled = false;
}

document.getElementById('btn-guardar').addEventListener('click', guardarCalificaciones);
document.getElementById('btn-guardar-bottom').addEventListener('click', guardarCalificaciones);