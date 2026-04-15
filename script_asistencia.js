// ════════════════════════════════════════════════════════
//  CONFIGURACIÓN
// ════════════════════════════════════════════════════════
const API_KEY           = "AIzaSyA1GsdvskVoIQDSOg6EI8aEQW5MnQsFeoE";
const SPREADSHEET_ID    = "1ytP_bXfbjRsA05KJYM8HAxQvTvnAeHgz09r4QhMy-Rs";
const SHEET_ESTUDIANTES = "asignaturas_estudiantes";
const SHEET_PERIODOS    = "Periodos";
const SHEET_ASISTENCIA  = "Asistencia";

const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbygBydUcKQWczkcs4jL3BshWNY1npLtsqK9IbNA72nIgpRhmjHo3tKnCVry0AcwvGIf/exec";

// ════════════════════════════════════════════════════════
//  ESTADO GLOBAL
// ════════════════════════════════════════════════════════
let nombreDocente       = "";
let sedeDocente         = "";
let todosLosEstudiantes = [];
let headersEstudiantes  = [];
let estudiantesActuales = [];
let asignaturasDocente  = [];
let gradoActivo         = "";
let periodoActivo       = "";
let periodosBloqueados  = {};
let fallasPrevias       = {};
let datosListos         = false;  // ← bandera: indica que inicializar() terminó

// Matriz de inputs para navegación con flechas [fila][col]
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

function fetchSheet(sheetName) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(sheetName)}?key=${API_KEY}`;
  return fetch(url).then(r => r.json());
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

function setSaveStatus(msg, tipo) {
  const el = document.getElementById('save-status');
  el.textContent = msg;
  el.className = 'save-status' + (tipo ? ' ' + tipo : '');
}

function mostrarEstado(id) {
  ['state-loading','state-empty','table-section'].forEach(s => {
    document.getElementById(s).style.display = 'none';
  });
  if (id) {
    document.getElementById(id).style.display = id === 'table-section' ? 'block' : 'flex';
  }
}

const METADATOS = ['no_identificador','nombres_apellidos','nombres y apellidos','nombres y apeliidos','grado','sede','estado','comportamiento','id_estudiante','contrasena','no documento','no_documento','documento','edad','fecha_nacimiento','fecha nacimiento','acudiente','telefono','direccion','correo','eps','genero','sexo','rh','tipo_documento','tipo documento','jornada','foto_url'];
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
//  CARGA INICIAL
//  Al terminar, si el docente ya había elegido el periodo,
//  se dispara el filtro de grados automáticamente.
// ════════════════════════════════════════════════════════
async function inicializar() {
  document.getElementById('loader-grado').classList.add('visible');
  try {
    const [dataEst, dataPer] = await Promise.all([
      fetchSheet(SHEET_ESTUDIANTES),
      fetchSheet(SHEET_PERIODOS)
    ]);

    // Procesar periodos → determinar bloqueos
    if (!dataPer.error && dataPer.values && dataPer.values.length > 1) {
      const hPer    = dataPer.values[0].map(h => (h||'').trim());
      const iNum    = hPer.findIndex(h => h.toUpperCase().includes('PERIODO'));
      const iInicio = hPer.findIndex(h => h.toUpperCase().includes('INICIO'));
      const iCierre = hPer.findIndex(h =>
        h.toUpperCase().includes('CIERRE') && h.toUpperCase().includes('SISTEMA')
      );
      const hoy = new Date(); hoy.setHours(0,0,0,0);
      dataPer.values.slice(1).forEach(row => {
        const num    = String(row[iNum]    || '').trim();
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
    if (!dataEst.error && dataEst.values && dataEst.values.length > 1) {
      headersEstudiantes  = dataEst.values[0].map(h => (h||'').trim());
      todosLosEstudiantes = dataEst.values.slice(1);
    }

    // Marcar datos como listos
    datosListos = true;

  } catch(err) {
    console.error('Error inicializando:', err);
  } finally {
    document.getElementById('loader-grado').classList.remove('visible');

    // Si el docente ya eligió periodo antes de que terminara la carga,
    // disparar el filtro ahora que los datos están disponibles.
    const selPer = document.getElementById('sel-periodo');
    if (selPer.value) {
      poblarGrados(selPer.value);
    }
  }
}
inicializar();

// ════════════════════════════════════════════════════════
//  PASO 1 — Periodo → poblar grados
//  El listener solo llama a poblarGrados().
//  Si los datos aún no están listos, muestra el loader
//  y espera a que inicializar() los active.
// ════════════════════════════════════════════════════════
document.getElementById('sel-periodo').addEventListener('change', function () {
  const periodo = this.value;

  // Resetear selector de grado siempre
  const selGrado = document.getElementById('sel-grado');
  selGrado.innerHTML = '<option value="">— Elige un grado —</option>';
  selGrado.disabled  = true;
  mostrarEstado(null);
  document.getElementById('selection-summary').style.display = 'none';

  if (!periodo) return;

  if (!datosListos) {
    // Datos aún cargando: mostrar loader y esperar.
    // El finally de inicializar() se encargará de disparar poblarGrados.
    document.getElementById('loader-grado').classList.add('visible');
    return;
  }

  poblarGrados(periodo);
});

// ════════════════════════════════════════════════════════
//  POBLAR GRADOS — lógica separada del listener
// ════════════════════════════════════════════════════════
function poblarGrados(periodo) {
  const selGrado = document.getElementById('sel-grado');

  if (!todosLosEstudiantes.length) {
    selGrado.innerHTML = '<option value="">Error al cargar datos</option>';
    return;
  }

  const iGrado   = headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO');
  const iEstado  = headersEstudiantes.findIndex(h => h.toUpperCase() === 'ESTADO');
  const colsAsig = headersEstudiantes
    .map((h, i) => ({ nombre: h, idx: i }))
    .filter(c => !METADATOS.includes(c.nombre.toLowerCase().trim()));

  const filasDelDocente = todosLosEstudiantes.filter(row => {
    const estado = (row[iEstado] || '').trim().toLowerCase();
    if (estado === 'inactivo' || estado === 'retirado') return false;
    return colsAsig.some(c => (row[c.idx] || '').trim() === nombreDocente.trim());
  });

  const gradosUnicos = ordenarGrados(
    [...new Set(filasDelDocente.map(r => (r[iGrado] || '').trim()))].filter(Boolean)
  );

  if (!gradosUnicos.length) {
    selGrado.innerHTML = '<option value="">Sin grados asignados</option>';
    return;
  }

  selGrado.innerHTML =
    '<option value="">— Elige un grado —</option>' +    
    gradosUnicos.map(g => `<option value="${g}">${g}</option>`).join('');
  selGrado.disabled = false;
}

// ════════════════════════════════════════════════════════
//  PASO 2 — Grado → cargar tabla
// ════════════════════════════════════════════════════════
document.getElementById('sel-grado').addEventListener('change', function () {
  gradoActivo   = this.value;
  periodoActivo = document.getElementById('sel-periodo').value;
  mostrarEstado(null);
  document.getElementById('selection-summary').style.display = 'none';

  if (gradoActivo) cargarAsistencia();
});

// ════════════════════════════════════════════════════════
//  CARGAR TABLA DE ASISTENCIA
// ════════════════════════════════════════════════════════
async function cargarAsistencia() {
  mostrarEstado('state-loading');
  document.getElementById('state-loading').querySelector('p').textContent = 'Cargando estudiantes...';

  const iGrado   = headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO');
  const iEstado  = headersEstudiantes.findIndex(h => h.toUpperCase() === 'ESTADO');
  const colsAsig = headersEstudiantes
    .map((h, i) => ({ nombre: h, idx: i }))
    .filter(c => !METADATOS.includes(c.nombre.toLowerCase().trim()));

  // Solo estudiantes del grado que además tengan al docente asignado
  const filasFiltradas = todosLosEstudiantes.filter(row => {
    const estadoFila = (row[iEstado] || '').trim().toLowerCase();
    if (estadoFila === 'inactivo' || estadoFila === 'retirado') return false;
    const gradoOk = gradoActivo === '__todos__'
      ? true
      : (row[iGrado] || '').trim().toLowerCase() === gradoActivo.toLowerCase();
    const tieneDocente = colsAsig.some(c => (row[c.idx] || '').trim() === nombreDocente.trim());
    return gradoOk && tieneDocente;
  });

  if (gradoActivo === '__todos__') {
    const iNombreSort = headersEstudiantes.findIndex(h =>
      h.trim().toLowerCase() === 'nombres y apellidos' ||
      h.trim().toLowerCase() === 'nombres_apellidos'
    );
    filasFiltradas.sort((a, b) => {
      const ORDEN = ['Preescolar','Primero','Segundo','Tercero','Cuarto','Quinto',
                     'Sexto','Séptimo','Octavo','Noveno','Décimo','Undécimo'];
      const ia = ORDEN.indexOf((a[iGrado]||'').trim());
      const ib = ORDEN.indexOf((b[iGrado]||'').trim());
      if (ia !== ib) return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
      return (a[iNombreSort]||'').localeCompare(b[iNombreSort]||'');
    });
  }

  if (!filasFiltradas.length) {
    mostrarEstado(null);
    document.getElementById('state-empty-msg').textContent =
      `No tienes estudiantes asignados en ${gradoActivo}.`;
    mostrarEstado('state-empty');
    return;
  }

  estudiantesActuales = filasFiltradas;

  asignaturasDocente = colsAsig.filter(c =>
    filasFiltradas.some(row => (row[c.idx] || '').trim() === nombreDocente.trim())
  );

  if (!asignaturasDocente.length) {
    mostrarEstado(null);
    document.getElementById('state-empty-msg').textContent =
      `No tienes asignaturas asignadas en ${gradoActivo}.`;
    mostrarEstado('state-empty');
    return;
  }

  document.getElementById('state-loading').querySelector('p').textContent = 'Recuperando registros anteriores...';
  fallasPrevias = await cargarFallasPrevias();

  mostrarEstado(null);
  renderTabla();

  document.getElementById('chip-periodo').textContent = `Periodo ${periodoActivo}`;
  document.getElementById('chip-grado').textContent =
    gradoActivo === '__todos__' ? 'Todos los grados' : gradoActivo;
  document.getElementById('selection-summary').style.display = 'flex';
}

// ════════════════════════════════════════════════════════
//  CARGAR FALLAS PREVIAS
// ════════════════════════════════════════════════════════
async function cargarFallasPrevias() {
  try {
    const data = await fetchSheet(SHEET_ASISTENCIA);
    if (data.error || !data.values || data.values.length < 2) return {};

    const headers = data.values[0].map(h => (h||'').trim());
    const rows    = data.values.slice(1);

    const iP    = headers.findIndex(h => h.toUpperCase() === 'PERIODO');
    const iG    = headers.findIndex(h => h.toUpperCase() === 'GRADO');
    const iId   = headers.findIndex(h => h.toUpperCase() === 'ID_ESTUDIANTE');
    const iAsig = headers.findIndex(h => h.toUpperCase() === 'ASIGNATURA');
    const iF    = headers.findIndex(h => h.toUpperCase() === 'FALLAS');
    const iNom  = headers.findIndex(h =>
      h.toUpperCase() === 'NOMBRE' ||
      h.toUpperCase() === 'NOMBRES Y APELLIDOS' ||
      h.toUpperCase() === 'NOMBRES_APELLIDOS'
    );

    const resultado = {};
    rows.forEach(row => {
      // Comparar periodo normalizado (como string, removiendo .0 final)
      const perRow = String(row[iP]||'').trim().replace(/\.0$/, '');
      const perAct = String(periodoActivo).trim().replace(/\.0$/, '');
      if (perRow !== perAct) return;
      if (gradoActivo !== '__todos__' &&
          (row[iG]||'').trim().toLowerCase() !== gradoActivo.toLowerCase()) return;
      const id   = iId >= 0 ? String(row[iId]||'').trim() : '';
      const nombre = iNom >= 0 ? String(row[iNom]||'').trim() : '';
      const asig = String(row[iAsig]||'').trim();
      const f    = String(row[iF]||'').trim();
      if (asig && f !== '') {
        // Guardar por ID (si existe)
        if (id) {
          if (!resultado[id]) resultado[id] = {};
          resultado[id][asig] = f;
        }
        // También guardar por nombre como fallback
        if (nombre) {
          if (!resultado[nombre]) resultado[nombre] = {};
          resultado[nombre][asig] = f;
        }
      }
    });
    return resultado;
  } catch(err) {
    console.warn('No se pudieron cargar fallas previas:', err);
    return {};
  }
}

// ════════════════════════════════════════════════════════
//  RENDERIZAR TABLA con navegación por flechas
// ════════════════════════════════════════════════════════
function renderTabla() {
  const iNombre = headersEstudiantes.findIndex(h =>
    h.trim().toLowerCase() === 'nombres y apellidos' ||
    h.trim().toLowerCase() === 'nombres_apellidos'   ||
    h.trim().toLowerCase() === 'nombres y apeliidos'
  );
  const iId  = headersEstudiantes.findIndex(h =>
    h.toUpperCase() === 'ID_ESTUDIANTE' ||
    h.toUpperCase() === 'NO DOCUMENTO' ||
    h.toUpperCase() === 'NO_DOCUMENTO' ||
    h.toUpperCase() === 'DOCUMENTO'
  );
  const iSeq  = headersEstudiantes.findIndex(h => h.toUpperCase() === 'NO_IDENTIFICADOR');
  const iSede = headersEstudiantes.findIndex(h => h.toUpperCase() === 'SEDE');

  const mostrarGrado = gradoActivo === '__todos__';
  const thGradoExist = document.getElementById('th-grado-asist');
  if (mostrarGrado && !thGradoExist) {
    const thG = document.createElement('th');
    thG.id = 'th-grado-asist'; thG.className = 'th-fixed';
    thG.textContent = 'Grado';
    const theadRef = document.getElementById('tabla-headers');
    theadRef.insertBefore(thG, theadRef.children[3]);
  } else if (!mostrarGrado && thGradoExist) {
    thGradoExist.remove();
  }

  // Encabezados dinámicos
  const theadRow = document.getElementById('tabla-headers');
  while (theadRow.children.length > 3) theadRow.removeChild(theadRow.lastChild);
  asignaturasDocente.forEach(asig => {
    const th = document.createElement('th');
    th.className   = 'th-asig';
    th.textContent = asig.nombre;
    theadRow.appendChild(th);
  });

  const tbody = document.getElementById('tabla-body');
  tbody.innerHTML = '';
  matrizInputs    = [];

  estudiantesActuales.forEach((row, rowIdx) => {
    const nombre  = (iNombre >= 0 ? row[iNombre] : '') || 'Sin nombre';
    const sede    = (row[iSede]  || '—').trim();
    const idEstud = (row[iId]    || '').trim();
    const nombreTrim = nombre.trim();
    // Buscar fallas previas: primero por ID, luego por nombre completo
    const prev    = (idEstud && fallasPrevias[idEstud]) ? fallasPrevias[idEstud]
                  : (fallasPrevias[nombreTrim] || {});
    const tienePrev = Object.keys(prev).length > 0;

    const tr = document.createElement('tr');
    tr.dataset.rowIdx  = rowIdx;
    tr.dataset.idEstud = idEstud;
    if (tienePrev) tr.classList.add('fila-con-previas');

    const gradoFila = (row[headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO')] || '').trim();
    tr.innerHTML = `
      <td class="td-num">${rowIdx + 1}</td>
      <td class="td-nombre ${tienePrev ? 'tiene-datos' : ''}" title="${nombre.trim()}">${nombre.trim()}</td>
      <td class="td-sede">${sede}</td>
      ${mostrarGrado ? `<td class="td-sede" style="color:var(--iean-dark);font-weight:600;">${gradoFila}</td>` : ''}
    `;

    const filaInputs = [];

    asignaturasDocente.forEach((asig, colIdx) => {
      // Buscar valor previo con tolerancia: exacto, luego case-insensitive
      let valorPrev = prev[asig.nombre] || '';
      if (!valorPrev) {
        // Buscar por nombre normalizado (case-insensitive, sin espacios extra)
        const asigNorm = asig.nombre.trim().toLowerCase();
        const keyMatch = Object.keys(prev).find(k => k.trim().toLowerCase() === asigNorm);
        if (keyMatch) valorPrev = prev[keyMatch];
      }
      const td    = document.createElement('td');
      td.className = 'td-falla';

      const input = document.createElement('input');
      input.type         = 'number';
      input.className    = 'falla-input';
      input.min          = '0';
      input.max          = '99';
      input.step         = '1';
      input.placeholder  = '—';
      input.value = valorPrev !== '' ? valorPrev : '';
      input.dataset.asig = asig.nombre;
      input.dataset.row  = rowIdx;
      input.dataset.col  = colIdx;

      if (valorPrev && parseInt(valorPrev) > 0) input.classList.add('con-fallas');

      input.addEventListener('input', () => {
        validarFalla(input);
        setSaveStatus('');
        const val = parseInt(input.value);
        if (!isNaN(val) && val > 0) {
          input.classList.add('con-fallas');
          tr.classList.add('fila-con-previas');
          tr.querySelector('.td-nombre').classList.add('tiene-datos');
        } else {
          input.classList.remove('con-fallas');
        }
      });

      input.addEventListener('keydown', e => manejarTeclado(e, rowIdx, colIdx));

      td.appendChild(input);
      tr.appendChild(td);
      filaInputs.push(input);
    });

    matrizInputs.push(filaInputs);
    tbody.appendChild(tr);
  });

  document.getElementById('table-title').textContent =
    gradoActivo === '__todos__' ? 'Fallas — Todos los grados' : `Fallas — Grado ${gradoActivo}`;
  document.getElementById('table-count').textContent =
    `${estudiantesActuales.length} estudiante${estudiantesActuales.length !== 1 ? 's' : ''}` +
    ` · ${asignaturasDocente.length} asignatura${asignaturasDocente.length !== 1 ? 's' : ''}`;

  mostrarEstado('table-section');

  const bloqueado = periodosBloqueados[periodoActivo] === true;
  if (bloqueado) {
    bloquearTabla();
    setSaveStatus('🔒 El sistema no está habilitado para registrar fallas en este periodo.', 'error');
  } else {
    const tienePrevias = Object.keys(fallasPrevias).length > 0;
    setSaveStatus(tienePrevias ? '📋 Fallas anteriores cargadas. Edita y guarda cuando estés listo.' : '');
    if (matrizInputs.length > 0 && matrizInputs[0].length > 0) {
      setTimeout(() => matrizInputs[0][0].focus(), 100);
    }
  }
}

// ════════════════════════════════════════════════════════
//  NAVEGACIÓN CON TECLADO
// ════════════════════════════════════════════════════════
function manejarTeclado(e, fila, col) {
  const totalFilas = matrizInputs.length;
  const totalCols  = asignaturasDocente.length;
  let nuevaFila = fila, nuevaCol = col;

  switch (e.key) {
    case 'ArrowRight':
      if (col < totalCols - 1) nuevaCol = col + 1;
      else if (fila < totalFilas - 1) { nuevaFila = fila + 1; nuevaCol = 0; }
      e.preventDefault(); break;
    case 'ArrowLeft':
      if (col > 0) nuevaCol = col - 1;
      else if (fila > 0) { nuevaFila = fila - 1; nuevaCol = totalCols - 1; }
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
        e.preventDefault(); nuevaFila = fila - 1; nuevaCol = totalCols - 1;
      } else return;
      break;
    default: return;
  }

  if (matrizInputs[nuevaFila] && matrizInputs[nuevaFila][nuevaCol]) {
    const dest = matrizInputs[nuevaFila][nuevaCol];
    dest.focus();
    dest.select();
  }
}

// ════════════════════════════════════════════════════════
//  BLOQUEAR TABLA
// ════════════════════════════════════════════════════════
function bloquearTabla() {
  document.querySelectorAll('.falla-input').forEach(input => {
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
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
        <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
      </svg>
      Periodo no disponible — Modo de lectura
    `;
    headerBar.appendChild(banner);
  }
}

// ════════════════════════════════════════════════════════
//  VALIDAR INPUT
// ════════════════════════════════════════════════════════
function validarFalla(input) {
  if (input.value === '') { input.classList.remove('invalid'); return true; }
  const v  = parseInt(input.value);
  const ok = !isNaN(v) && v >= 0 && v <= 99;
  input.classList.toggle('invalid', !ok);
  return ok;
}

// ════════════════════════════════════════════════════════
//  LIMPIAR
// ════════════════════════════════════════════════════════
document.getElementById('btn-limpiar').addEventListener('click', () => {
  if (!confirm('¿Deseas limpiar todos los campos? Los datos no guardados se perderán.')) return;
  document.querySelectorAll('.falla-input').forEach(i => {
    i.value = '';
    i.classList.remove('invalid', 'con-fallas');
  });
  document.querySelectorAll('.td-nombre').forEach(c => c.classList.remove('tiene-datos'));
  document.querySelectorAll('tr').forEach(r => r.classList.remove('fila-con-previas'));
  setSaveStatus('');
});

// ════════════════════════════════════════════════════════
//  RECOPILAR DATOS
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
  const fecha     = new Date().toLocaleDateString('es-CO');

  document.querySelectorAll('#tabla-body tr').forEach((tr, idx) => {
    const row = estudiantesActuales[idx];
    if (!row) return;
    const id     = (row[iId]     || '').trim();
    const nombre = (iNombre >= 0 ? (row[iNombre]||'') : '').trim();
    const sede   = (row[iSede]   || '').trim();

    tr.querySelectorAll('.falla-input').forEach(input => {
      const asig   = input.dataset.asig;
      const fallas = input.value.trim();
      if (fallas !== '') {
        registros.push({
          periodo: periodoActivo, fecha,
          docente: nombreDocente,
          grado: gradoActivo === '__todos__'
            ? (row[headersEstudiantes.findIndex(h => h.toUpperCase() === 'GRADO')] || '').trim()
            : gradoActivo,
          sede, asignatura: asig,
          id:    (row[iId]  || '').trim(),
          idSeq: (row[iSeq] || '').trim(),
          nombre,
          fallas: parseInt(fallas) || 0
        });
      }
    });
  });
  return registros;
}

// ════════════════════════════════════════════════════════
//  ENVÍO A APPS SCRIPT
// ════════════════════════════════════════════════════════
function enviarAAppsScript(registros) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify({ tipo: 'asistencia', registros });
    fetch(APPS_SCRIPT_URL, {
      method: 'POST', headers: { 'Content-Type': 'text/plain' },
      body: payload, redirect: 'follow'
    })
    .then(r => r.text().then(text => {
      try { resolve(JSON.parse(text)); }
      catch { resolve({ status:'ok', guardados: registros.length }); }
    }))
    .catch(() => enviarConIframe(payload).then(resolve).catch(reject));
  });
}

function enviarConIframe(payload) {
  return new Promise(resolve => {
    const iframeName = 'sianario_asist_' + Date.now();
    const iframe = document.createElement('iframe');
    iframe.name = iframeName; iframe.style.display = 'none';
    document.body.appendChild(iframe);

    const form = document.createElement('form');
    form.method = 'POST'; form.action = APPS_SCRIPT_URL;
    form.target = iframeName; form.enctype = 'application/x-www-form-urlencoded';

    const inp = document.createElement('input');
    inp.type = 'hidden'; inp.name = 'payload'; inp.value = payload;
    form.appendChild(inp); document.body.appendChild(form);

    const limpiar = () => {
      try { document.body.removeChild(iframe); } catch(e) {}
      try { document.body.removeChild(form);   } catch(e) {}
    };

    const timeout = setTimeout(() => {
      limpiar();
      resolve({ status:'ok', mensaje:'Datos enviados. Confirma en tu Sheets.' });
    }, 12000);

    iframe.onload = () => {
      clearTimeout(timeout);
      try { resolve(JSON.parse(iframe.contentDocument?.body?.innerText||'')); }
      catch { resolve({ status:'ok', mensaje:'Datos enviados.' }); }
      limpiar();
    };
    form.submit();
  });
}

// ════════════════════════════════════════════════════════
//  GUARDAR
// ════════════════════════════════════════════════════════
async function guardarAsistencia() {
  if (periodosBloqueados[periodoActivo] === true) {
    setSaveStatus('🔒 No se puede guardar: el periodo no está disponible.', 'error');
    return;
  }
  const invalidos = document.querySelectorAll('.falla-input.invalid');
  if (invalidos.length > 0) {
    setSaveStatus('Corrige los valores en rojo antes de guardar.', 'error');
    invalidos[0].focus(); return;
  }
  const registros = recopilarDatos();
  if (!registros.length) {
    setSaveStatus('No hay datos para guardar. Ingresa al menos un valor.', 'error');
    return;
  }

  const btnTop = document.getElementById('btn-guardar');
  const btnBot = document.getElementById('btn-guardar-bottom');
  btnTop.disabled = btnBot.disabled = true;
  setSaveStatus('Guardando Asistencia...', '');

  try {
    const result = await enviarAAppsScript(registros);
    if (result.status === 'ok') {
      setSaveStatus(`✓ ${result.guardados || registros.length} registro(s) guardados correctamente.`, 'success');
    } else if (result.status === 'parcial') {
      setSaveStatus(`⚠️ Guardado parcial: ${result.mensaje}`, 'error');
    } else {
      setSaveStatus(`Error: ${result.mensaje}`, 'error');
    }
  } catch(err) {
    setSaveStatus('Error al enviar datos. Verifica tu conexión.', 'error');
    console.error(err);
  }
  btnTop.disabled = btnBot.disabled = false;
}

document.getElementById('btn-guardar').addEventListener('click', guardarAsistencia);
document.getElementById('btn-guardar-bottom').addEventListener('click', guardarAsistencia);