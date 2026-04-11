// ════════════════════════════════════════════════════════
//  CONFIGURACIÓN
// ════════════════════════════════════════════════════════
const API_KEY        = "AIzaSyA1GsdvskVoIQDSOg6EI8aEQW5MnQsFeoE";
const SPREADSHEET_ID = "1ytP_bXfbjRsA05KJYM8HAxQvTvnAeHgz09r4QhMy-Rs";

const SHEET_ESTUDIANTES   = "asignaturas_estudiantes";
const SHEET_PERIODOS      = "Periodos";
const SHEET_AVISOS        = "Avisos";
const SHEET_CALIFICACIONES = "Calificaciones";

const MODULOS = {
  calificaciones: 'calificaciones.html',
  indicadores:    'indicadores.html',
  asistencia:     'asistencia.html'
  // 'observaciones' se maneja como panel en página, no como redirección
};

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

function norm(h) {
  return String(h || '').trim().toUpperCase()
    .replace(/[ÁÀÂÄ]/g,'A').replace(/[ÉÈÊË]/g,'E').replace(/[ÍÌÎÏ]/g,'I')
    .replace(/[ÓÒÔÖ]/g,'O').replace(/[ÚÙÛÜ]/g,'U')
    .replace(/[áàâä]/g,'A').replace(/[éèêë]/g,'E').replace(/[íìîï]/g,'I')
    .replace(/[óòôö]/g,'O').replace(/[úùûü]/g,'U')
    .replace(/[ñÑ]/g,'N');
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

function formatFecha(d) {
  if (!d) return '';
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}

// ════════════════════════════════════════════════════════
//  PERFIL
// ════════════════════════════════════════════════════════
const nombreDocente = getUrlParameter('nombre');
const sedeDocente   = getUrlParameter('sede');

document.getElementById('nombre-docente').textContent = nombreDocente || 'Docente';
document.getElementById('perfil-sede').textContent    = sedeDocente   || 'Sede no asignada';

// ════════════════════════════════════════════════════════
//  LOGOUT
// ════════════════════════════════════════════════════════
document.getElementById('btn-logout').addEventListener('click', () => {
  if (confirm('¿Estás seguro de que deseas cerrar sesión?')) {
    window.location.href = 'Index.html';
  }
});

// ════════════════════════════════════════════════════════
//  MÓDULOS
// ════════════════════════════════════════════════════════
document.querySelectorAll('.action-card').forEach(card => {
  card.addEventListener('click', function () {
    const modulo  = this.dataset.modulo;
    if (modulo === 'observaciones') {
      abrirPanelObservaciones();
      return;
    }
    const destino = MODULOS[modulo];
    if (destino) {
      const params = `?nombre=${encodeURIComponent(nombreDocente)}&sede=${encodeURIComponent(sedeDocente)}`;
      window.location.href = destino + params;
    } else {
      alert('Módulo "' + modulo + '" aún no está disponible.');
    }
  });
});

// ════════════════════════════════════════════════════════
//  INICIALIZACIÓN PRINCIPAL
//  Carga en paralelo: estudiantes, periodos, avisos y calificaciones
// ════════════════════════════════════════════════════════
async function cargarDashboard() {
  try {
    const [dataEst, dataPer, dataAvisos, dataCalif] = await Promise.all([
      fetchSheet(SHEET_ESTUDIANTES),
      fetchSheet(SHEET_PERIODOS),
      fetchSheet(SHEET_AVISOS),
      fetchSheet(SHEET_CALIFICACIONES)
    ]);

    // 1. Stats
    const periodosData = procesarPeriodos(dataPer.values || []);
    if (!dataEst.error && dataEst.values?.length > 1) {
      procesarEstudiantes(dataEst.values);
    } else {
      setStatError();
    }

    // 2. Avisos: combina avisos del admin + notificaciones automáticas de cierre
    const avisosAdmin = extraerAvisosAdmin(dataAvisos);
    const avisosAuto  = generarAvisosAutomaticos(periodosData);
    renderAvisos([...avisosAuto, ...avisosAdmin]);

    // 3. Notas bajas
    if (!dataCalif.error && dataCalif.values?.length > 1 &&
        !dataEst.error && dataEst.values?.length > 1) {
      procesarNotasBajas(dataCalif.values, dataEst.values);
    }

  } catch(err) {
    console.error('Error cargando dashboard:', err);
    setStatError();
  }
}

// ════════════════════════════════════════════════════════
//  STATS — ESTUDIANTES
// ════════════════════════════════════════════════════════
function procesarEstudiantes(values) {
  const headers  = values[0].map(h => (h || '').trim());
  const rows     = values.slice(1);

  const iEstado  = headers.findIndex(h => norm(h) === 'ESTADO');
  const iGrado   = headers.findIndex(h => norm(h) === 'GRADO');
  const iComport = headers.findIndex(h => norm(h) === 'COMPORTAMIENTO');

  const metadatos = ['no_identificador','nombres_apellidos','grado','sede','estado','comportamiento'];
  const colsAsig  = headers
    .map((h, i) => ({ nombre: h, idx: i }))
    .filter(c => !metadatos.includes(c.nombre.toLowerCase().trim()));

  const activas = rows.filter(row => {
    const estado = (row[iEstado] || '').trim().toLowerCase();
    return estado !== 'inactivo' && estado !== 'retirado';
  });

  // Dirección de grado
  let gradoDireccion = null;
  if (iComport >= 0) {
    const dirs = activas.filter(row => (row[iComport] || '').trim() === nombreDocente.trim());
    if (dirs.length) {
      const grados = [...new Set(dirs.map(r => (r[iGrado] || '').trim()))].filter(Boolean);
      gradoDireccion = grados.join(', ');
    }
  }

  const elDir    = document.getElementById('stat-direccion');
  const elDirSub = document.getElementById('stat-direccion-sub');
  if (gradoDireccion) {
    elDir.textContent    = gradoDireccion;
    elDirSub.textContent = '';
  } else {
    elDir.textContent    = 'N/A';
    elDirSub.textContent = 'sin dirección asignada';
  }

  // Estudiantes activos a cargo
  const total = activas.filter(row =>
    colsAsig.some(c => (row[c.idx] || '').trim() === nombreDocente.trim())
  ).length;
  document.getElementById('stat-estudiantes').textContent = total || '0';
  document.getElementById('stat-anio').textContent = new Date().getFullYear();
}

function setStatError() {
  ['stat-direccion','stat-estudiantes','stat-periodo'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = 'N/A';
  });
  const sub = document.getElementById('stat-periodo-sub');
  if (sub) { sub.textContent = 'sin conexión'; }
  const dirSub = document.getElementById('stat-direccion-sub');
  if (dirSub) dirSub.textContent = 'sin conexión';
}

// ════════════════════════════════════════════════════════
//  STATS — PERIODOS
//  Devuelve array de periodos procesados para reusar luego
// ════════════════════════════════════════════════════════
function procesarPeriodos(values) {
  if (!values || values.length < 2) return [];

  const headers  = values[0].map(h => (h || '').trim());
  const rows     = values.slice(1);

  const iPeriodo    = headers.findIndex(h => norm(h).includes('PERIODO'));
  const iInicio     = headers.findIndex(h => norm(h).includes('INICIO'));
  const iFinal      = headers.findIndex(h =>
    (norm(h).includes('FINAL') || norm(h).includes('FIN')) &&
    !norm(h).includes('CIERRE') && !norm(h).includes('SISTEMA')
  );
  const iCieSistema = headers.findIndex(h =>
    norm(h).includes('CIERRE') && norm(h).includes('SISTEMA')
  );

  const hoy = new Date(); hoy.setHours(0,0,0,0);

  const periodosProc = rows.map(row => ({
    numero:        String(row[iPeriodo] || '').trim(),
    inicio:        parseFecha(row[iInicio]),
    fin:           iFinal       >= 0 ? parseFecha(row[iFinal])      : null,
    cierreSistema: iCieSistema  >= 0 ? parseFecha(row[iCieSistema]) : null
  })).filter(p => p.numero);

  // Encontrar periodo activo o próximo
  let periodoActivo = null;
  for (const p of periodosProc) {
    if (!p.inicio) continue;
    const cierreEfect = p.cierreSistema || p.fin;
    if (hoy >= p.inicio && (!cierreEfect || hoy <= cierreEfect)) {
      periodoActivo = p;
      break;
    }
  }

  const elP    = document.getElementById('stat-periodo');
  const elPSub = document.getElementById('stat-periodo-sub');

  if (periodoActivo) {
    elP.textContent = periodoActivo.numero;
    const cierreEfect = periodoActivo.cierreSistema || periodoActivo.fin;
    if (cierreEfect) {
      const dias = Math.ceil((cierreEfect - hoy) / 86400000);
      if (dias <= 0) {
        elPSub.textContent = 'Sistema cerrado';
        elPSub.style.color = '#b71c1c';
      } else {
        elPSub.textContent = `Cierra en ${dias} día${dias !== 1 ? 's' : ''}`;
        elPSub.style.color = dias <= 5 ? '#e65100' : '';
      }
    } else {
      elPSub.textContent = 'en curso';
      elPSub.style.color = '';
    }
  } else {
    const proximos = periodosProc.filter(p => p.inicio && p.inicio > hoy);
    if (proximos.length) {
      elP.textContent = proximos[0].numero;
      const dias = Math.ceil((proximos[0].inicio - hoy) / 86400000);
      elPSub.textContent = `Inicia en ${dias} día${dias !== 1 ? 's' : ''}`;
      elPSub.style.color = '#1565c0';
    } else {
      elP.textContent    = '—';
      elPSub.textContent = 'Año finalizado';
      elPSub.style.color = '';
    }
  }

  return periodosProc;
}

// ════════════════════════════════════════════════════════
//  AVISOS DEL ADMINISTRADOR (hoja Avisos)
// ════════════════════════════════════════════════════════
function extraerAvisosAdmin(dataAvisos) {
  if (!dataAvisos || dataAvisos.error || !dataAvisos.values || dataAvisos.values.length < 2) return [];

  const headers = dataAvisos.values[0].map(h => (h || '').trim());
  const rows    = dataAvisos.values.slice(1).reverse();

  const iTitulo = headers.findIndex(h => norm(h).includes('TITULO'));
  const iMensaje= headers.findIndex(h => norm(h).includes('MENSAJE'));
  const iFecha  = headers.findIndex(h => norm(h).includes('FECHA'));
  const iDest   = headers.findIndex(h => norm(h).includes('DESTINATARIO') || norm(h).includes('DIRIGIDO'));
  const iActivo = headers.findIndex(h => norm(h).includes('ACTIVO'));

  return rows
    .filter(row => {
      const activo = (row[iActivo] || 'Si').trim().toLowerCase();
      if (activo === 'no' || activo === 'false' || activo === '0') return false;
      const dest = norm(row[iDest] || '');
      if (!dest || dest === 'TODOS' || dest === 'DOCENTES') return true;
      return dest === norm(nombreDocente);
    })
    .slice(0, 5)
    .map(row => ({
      tipo:    'admin',
      titulo:  (row[iTitulo]  || 'Aviso').trim(),
      mensaje: (row[iMensaje] || '').trim(),
      fecha:   (row[iFecha]   || '').trim(),
      urgente: false
    }));
}

// ════════════════════════════════════════════════════════
//  AVISOS AUTOMÁTICOS DE CIERRE DEL SISTEMA
// ════════════════════════════════════════════════════════
function generarAvisosAutomaticos(periodosProc) {
  const avisosAuto = [];
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  const UMBRAL_DIAS = 10;

  periodosProc.forEach(p => {
    const cierreEfect = p.cierreSistema || p.fin;
    if (!cierreEfect || !p.inicio) return;
    if (hoy < p.inicio || hoy > cierreEfect) return;

    const dias = Math.ceil((cierreEfect - hoy) / 86400000);
    if (dias < 0) return;

    if (dias === 0) {
      avisosAuto.push({
        tipo:    'auto-cierre',
        titulo:  `¡Hoy cierra el sistema — Periodo ${p.numero}!`,
        mensaje: `Esta es la última oportunidad para ingresar calificaciones, fallas e indicadores del Periodo ${p.numero}. El sistema cierra hoy.`,
        fecha:   formatFecha(hoy),
        urgente: true
      });
    } else if (dias <= UMBRAL_DIAS) {
      avisosAuto.push({
        tipo:    'auto-cierre',
        titulo:  `Cierre del sistema en ${dias} día${dias !== 1 ? 's' : ''} — Periodo ${p.numero}`,
        mensaje: `El sistema cierra el ${formatFecha(cierreEfect)}. Recuerda registrar calificaciones, fallas e indicadores antes de esa fecha.`,
        fecha:   formatFecha(hoy),
        urgente: dias <= 3
      });
    }
  });

  return avisosAuto;
}

// ════════════════════════════════════════════════════════
//  RENDERIZAR AVISOS
// ════════════════════════════════════════════════════════
function renderAvisos(avisos) {
  const contenedor = document.getElementById('notificaciones-container');
  const vacio      = document.getElementById('notif-empty');

  if (!avisos.length) {
    if (vacio) vacio.style.display = 'flex';
    return;
  }

  if (vacio) vacio.style.display = 'none';

  avisos.forEach(av => {
    const item = document.createElement('div');
    item.className = 'notification-item' + (av.urgente ? ' urgente' : '');

    const dot = document.createElement('div');
    dot.className = av.tipo === 'auto-cierre'
      ? (av.urgente ? 'notif-dot dot-rojo' : 'notif-dot dot-naranja')
      : 'notif-dot';

    const body = document.createElement('div');
    body.style.flex = '1';

    const titulo = document.createElement('div');
    titulo.className   = 'notif-title';
    titulo.textContent = av.titulo;

    const sub = document.createElement('div');
    sub.className   = 'notif-sub';
    sub.textContent = av.mensaje;

    const fecha = document.createElement('div');
    fecha.className   = 'notif-fecha';
    fecha.textContent = av.fecha;

    body.appendChild(titulo);
    body.appendChild(sub);
    if (av.fecha) body.appendChild(fecha);

    item.appendChild(dot);
    item.appendChild(body);
    contenedor.appendChild(item);
  });
}

// ════════════════════════════════════════════════════════
//  NOTAS BAJAS — Procesar hoja Calificaciones
// ════════════════════════════════════════════════════════
function procesarNotasBajas(califValues, estValues) {
  const hCalif    = califValues[0].map(h => (h || '').trim());
  const rowsCalif = califValues.slice(1);

  const iP      = hCalif.findIndex(h => norm(h) === 'PERIODO');
  const iAsig   = hCalif.findIndex(h => norm(h) === 'ASIGNATURA');
  const iGrad   = hCalif.findIndex(h => norm(h) === 'GRADO');
  const iDoc    = hCalif.findIndex(h => norm(h) === 'DOCENTE');
  const iId     = hCalif.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
  const iProm   = hCalif.findIndex(h => norm(h).includes('PROMEDIO') || norm(h).includes('PROM'));
  const iDesemp = hCalif.findIndex(h => norm(h).includes('DESEMPE'));

  if (iP < 0 || iAsig < 0 || iDoc < 0 || iProm < 0) return;

  const hEst    = estValues[0].map(h => (h || '').trim());
  const rowsEst = estValues.slice(1);
  const iIdEst  = hEst.findIndex(h => norm(h) === 'NO_IDENTIFICADOR');
  const iNombre = hEst.findIndex(h =>
    norm(h).includes('NOMBRE') || norm(h).includes('APELLIDO')
  );
  const mapaNombres = {};
  if (iIdEst >= 0 && iNombre >= 0) {
    rowsEst.forEach(r => {
      const id = (r[iIdEst] || '').trim();
      if (id) mapaNombres[id] = (r[iNombre] || '').trim();
    });
  }

  const notas_bajas = [];
  rowsCalif.forEach(row => {
    const docFila = (row[iDoc] || '').trim();
    if (docFila !== nombreDocente.trim()) return;

    const prom = parseFloat(String(row[iProm] || '').replace(',','.'));
    if (isNaN(prom) || prom >= 3.0) return;

    const idEst   = (row[iId]   || '').trim();
    const nombre  = mapaNombres[idEst] || idEst || 'Sin nombre';
    const grado   = (row[iGrad] || '').trim();
    const asig    = (row[iAsig] || '').trim();
    const periodo = (row[iP]    || '').trim();
    const desemp  = iDesemp >= 0 ? (row[iDesemp] || '').trim() : 'BAJO';

    notas_bajas.push({ nombre, grado, asig, periodo, prom, desemp });
  });

  notas_bajas.sort((a, b) => a.prom - b.prom);

  const grados = [...new Set(notas_bajas.map(n => n.grado))].filter(Boolean).sort();
  const selGrado = document.getElementById('filtro-grado-bajas');
  grados.forEach(g => {
    const opt = document.createElement('option');
    opt.value       = g;
    opt.textContent = g;
    selGrado.appendChild(opt);
  });

  document.getElementById('seccion-notas-bajas').style.display = 'block';
  document.getElementById('alerta-loading').style.display = 'none';

  window._notasBajasData = notas_bajas;
  renderTablaNotasBajas(notas_bajas);

  selGrado.addEventListener('change', aplicarFiltrosNotasBajas);
  document.getElementById('filtro-periodo-bajas').addEventListener('change', aplicarFiltrosNotasBajas);
}

function aplicarFiltrosNotasBajas() {
  const gradoFiltro   = document.getElementById('filtro-grado-bajas').value;
  const periodoFiltro = document.getElementById('filtro-periodo-bajas').value;
  const data          = window._notasBajasData || [];

  const filtradas = data.filter(n => {
    const okGrado   = !gradoFiltro   || n.grado   === gradoFiltro;
    const okPeriodo = !periodoFiltro || n.periodo  === periodoFiltro;
    return okGrado && okPeriodo;
  });

  renderTablaNotasBajas(filtradas);
}

function renderTablaNotasBajas(datos) {
  const tablaWrap = document.getElementById('alerta-tabla-wrap');
  const vacia     = document.getElementById('alerta-vacia');
  const conteo    = document.getElementById('alerta-conteo');
  const tbody     = document.getElementById('tbody-notas-bajas');

  conteo.textContent = datos.length > 0
    ? `${datos.length} estudiante${datos.length !== 1 ? 's' : ''}`
    : '0';

  if (!datos.length) {
    tablaWrap.style.display = 'none';
    vacia.style.display     = 'flex';
    return;
  }

  tablaWrap.style.display = 'block';
  vacia.style.display     = 'none';
  tbody.innerHTML = '';

  datos.forEach(n => {
    const promNum = parseFloat(n.prom);
    const tr = document.createElement('tr');

    let colorProm = '#c62828';
    if (promNum >= 2.5) colorProm = '#e65100';

    tr.innerHTML = `
      <td class="td-nombre-alerta">${n.nombre}</td>
      <td>${n.grado}</td>
      <td>${n.asig}</td>
      <td class="td-centro">P${n.periodo}</td>
      <td class="td-centro">
        <span class="chip-prom-bajo" style="background:${promNum >= 2.5 ? '#fff3e0' : '#ffebee'};
              border-color:${promNum >= 2.5 ? '#ffcc80' : '#ef9a9a'};
              color:${colorProm};">
          ${promNum.toFixed(1)}
        </span>
      </td>
      <td class="td-centro">
        <span class="chip-desemp-bajo">BAJO</span>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// ════════════════════════════════════════════════════════
//  ARRANCAR
// ════════════════════════════════════════════════════════
cargarDashboard();

// ════════════════════════════════════════════════════════
//  MÓDULO OBSERVACIONES
// ════════════════════════════════════════════════════════
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbygBydUcKQWczkcs4jL3BshWNY1npLtsqK9IbNA72nIgpRhmjHo3tKnCVry0AcwvGIf/exec";
const SHEET_OBS       = "Observaciones_Boletines";

const METADATOS_OBS = ['no_identificador','nombres_apellidos','grado','sede','estado'];
const ORDEN_GRADOS_OBS = [
  'Preescolar','Primero','Segundo','Tercero','Cuarto','Quinto',
  'Sexto','Séptimo','Octavo','Noveno','Décimo','Undécimo'
];

// Estado del módulo
let obsEstado = {
  gradosAsig:          {},   // { grado: [asignaturas...] }
  estudiantesPorGrado: {},   // { grado: [{id, nombre, docId}] }
  estudianteFiltrado:  [],
  estSeleccionado:     null, // { id, nombre, docId }
  cargaInicial:        false,
  todasObs:            [],
};

function ordenarGradosObs(lista) {
  return [...lista].sort((a, b) => {
    const ia = ORDEN_GRADOS_OBS.indexOf(a);
    const ib = ORDEN_GRADOS_OBS.indexOf(b);
    if (ia === -1 && ib === -1) return a.localeCompare(b);
    if (ia === -1) return 1; if (ib === -1) return -1;
    return ia - ib;
  });
}

// ── Abrir panel ──────────────────────────────────────────
function abrirPanelObservaciones() {
  const overlay = document.getElementById('obs-overlay');
  overlay.style.display = 'flex';
  document.body.style.overflow = 'hidden';

  activarTab('nueva');

  if (!obsEstado.cargaInicial) {
    cargarDatosObservaciones();
  }
}

function cerrarPanelObservaciones() {
  document.getElementById('obs-overlay').style.display = 'none';
  document.body.style.overflow = '';
}

document.getElementById('obs-btn-close').addEventListener('click', cerrarPanelObservaciones);
document.getElementById('obs-overlay').addEventListener('click', function(e) {
  if (e.target === this) cerrarPanelObservaciones();
});

// ── Tabs ─────────────────────────────────────────────────
function activarTab(tab) {
  ['nueva','consultar'].forEach(t => {
    document.getElementById('obs-tab-' + t).classList.toggle('obs-tab-active', t === tab);
    document.getElementById('obs-content-' + t).style.display = t === tab ? 'block' : 'none';
  });
  if (tab === 'consultar' && obsEstado.cargaInicial) {
    cargarObservacionesConsulta();
  }
}

document.getElementById('obs-tab-nueva').addEventListener('click', () => activarTab('nueva'));
document.getElementById('obs-tab-consultar').addEventListener('click', () => activarTab('consultar'));

// ── Carga inicial: asignaturas_estudiantes + Periodos en paralelo ──
// MEJORA 1: periodos habilitados según fechas activas de la hoja Periodos
// MEJORA 2: docId = columna ID_Estudiante (documento de identidad real)
async function cargarDatosObservaciones() {
  document.getElementById('obs-loading').style.display = 'flex';
  document.getElementById('obs-form-wrap').style.display = 'none';

  try {
    const [data, dataPer] = await Promise.all([
      fetchSheet(SHEET_ESTUDIANTES),
      fetchSheet(SHEET_PERIODOS)
    ]);

    // ── MEJORA 1: Periodos habilitados según fechas ───────
    const hoy = new Date(); hoy.setHours(0,0,0,0);
    const periodosHabilitados = new Set();

    if (dataPer.values && dataPer.values.length > 1) {
      const hPer    = dataPer.values[0].map(h => (h || '').trim());
      const rowsPer = dataPer.values.slice(1);

      const iPNum    = hPer.findIndex(h => norm(h).includes('PERIODO'));
      const iInicio  = hPer.findIndex(h => norm(h).includes('INICIO'));
      const iCieSist = hPer.findIndex(h => norm(h).includes('CIERRE') && norm(h).includes('SISTEMA'));
      const iFin     = hPer.findIndex(h =>
        (norm(h).includes('FINAL') || norm(h).includes('FIN')) &&
        !norm(h).includes('CIERRE') && !norm(h).includes('SISTEMA')
      );

      rowsPer.forEach(row => {
        const num    = String(row[iPNum] || '').trim();
        const inicio = parseFecha(row[iInicio]);
        // Preferir cierre del sistema; si no existe, usar fecha fin del periodo
        const cierre = (iCieSist >= 0 && row[iCieSist])
          ? parseFecha(row[iCieSist])
          : (iFin >= 0 ? parseFecha(row[iFin]) : null);

        if (num && inicio && cierre && hoy >= inicio && hoy <= cierre) {
          periodosHabilitados.add(num);
        }
      });
    }

    // Poblar select de periodo
    const selPer = document.getElementById('obs-select-periodo');
    selPer.innerHTML = '<option value="">Seleccionar...</option>';
    ['1','2','3','4'].forEach(p => {
      const opt = document.createElement('option');
      opt.value = p;
      if (periodosHabilitados.has(p)) {
        opt.textContent = `Periodo ${p}`;
      } else {
        opt.textContent = `Periodo ${p} (cerrado)`;
        opt.disabled = true;
      }
      selPer.appendChild(opt);
    });

    if (periodosHabilitados.size === 0) {
      selPer.innerHTML = '<option value="">Sin periodos activos</option>';
    }

    // ── Cargar estudiantes ────────────────────────────────
    if (data.error || !data.values || data.values.length < 2) {
      mostrarObsFeedback('No se pudo cargar la información del docente.', 'err');
      document.getElementById('obs-loading').style.display = 'none';
      return;
    }

    const headers = data.values[0].map(h => (h || '').trim());
    const rows    = data.values.slice(1);

    const iId     = headers.findIndex(h => norm(h) === 'NO_IDENTIFICADOR');
    // MEJORA 2: columna ID_Estudiante = documento de identidad del estudiante
    const iDocId  = headers.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
    const iNombre = headers.findIndex(h => norm(h).includes('NOMBRE') || norm(h).includes('APELLIDO'));
    const iGrado  = headers.findIndex(h => norm(h) === 'GRADO');
    const iEstado = headers.findIndex(h => norm(h) === 'ESTADO');

    // Columnas de asignaturas (excluir metadatos)
    const colsAsig = headers
      .map((h, i) => ({ nombre: h, idx: i }))
      .filter(c => !METADATOS_OBS.includes(c.nombre.toLowerCase().trim()));

    // Filtrar activos
    const activos = rows.filter(row => {
      const est = (row[iEstado] || '').trim().toLowerCase();
      return est !== 'inactivo' && est !== 'retirado';
    });

    // Construir gradosAsig y estudiantesPorGrado
    const gradosAsig  = {};
    const estPorGrado = {};

    activos.forEach(row => {
      const grado  = (row[iGrado]  || '').trim();
      const id     = (row[iId]     || '').trim();
      const nombre = (row[iNombre] || '').trim();
      // MEJORA 2: leer el documento de identidad real del estudiante
      const docId  = iDocId >= 0 ? (row[iDocId] || '').trim() : '';
      if (!grado || !id) return;

      colsAsig.forEach(c => {
        if ((row[c.idx] || '').trim() === nombreDocente.trim()) {
          if (!gradosAsig[grado]) gradosAsig[grado] = new Set();
          gradosAsig[grado].add(c.nombre);
        }
      });

      if (!estPorGrado[grado]) estPorGrado[grado] = [];
      if (id && nombre) {
        const existe = estPorGrado[grado].find(e => e.id === id);
        // MEJORA 2: incluir docId en el objeto del estudiante
        if (!existe) estPorGrado[grado].push({ id, nombre, docId });
      }
    });

    // Filtrar solo grados donde el docente imparte
    const gradosDocente = Object.keys(gradosAsig);
    Object.keys(estPorGrado).forEach(g => {
      if (!gradosDocente.includes(g)) delete estPorGrado[g];
    });

    // Convertir sets a arrays
    Object.keys(gradosAsig).forEach(g => {
      gradosAsig[g] = [...gradosAsig[g]].sort();
    });

    obsEstado.gradosAsig         = gradosAsig;
    obsEstado.estudiantesPorGrado = estPorGrado;
    obsEstado.cargaInicial        = true;

    // Poblar select de grado
    const selGrado = document.getElementById('obs-select-grado');
    selGrado.innerHTML = '<option value="">Seleccionar...</option>';
    ordenarGradosObs(gradosDocente).forEach(g => {
      const opt = document.createElement('option');
      opt.value = opt.textContent = g;
      selGrado.appendChild(opt);
    });

    // Poblar filtros de consulta
    const filtGrado = document.getElementById('obs-filtro-grado');
    filtGrado.innerHTML = '<option value="">Todos los grados</option>';
    ordenarGradosObs(gradosDocente).forEach(g => {
      const opt = document.createElement('option');
      opt.value = opt.textContent = g;
      filtGrado.appendChild(opt);
    });

    document.getElementById('obs-loading').style.display = 'none';
    document.getElementById('obs-form-wrap').style.display = 'block';

  } catch(err) {
    console.error('Error cargando obs:', err);
    document.getElementById('obs-loading').style.display = 'none';
    mostrarObsFeedback('Error de conexión al cargar datos.', 'err');
  }
}

// ── Cambio de grado → poblar asignaturas ──────────────────
document.getElementById('obs-select-grado').addEventListener('change', function() {
  const grado   = this.value;
  const selAsig = document.getElementById('obs-select-asignatura');
  selAsig.innerHTML = '';

  resetEstudiante();
  document.getElementById('obs-field-estudiante').style.display = 'none';
  document.getElementById('obs-field-texto').style.display = 'none';
  document.getElementById('obs-action-wrap').style.display = 'none';

  if (!grado) {
    selAsig.innerHTML = '<option value="">— elige grado primero —</option>';
    return;
  }

  const asigs = obsEstado.gradosAsig[grado] || [];
  selAsig.innerHTML = '<option value="">Seleccionar asignatura...</option>';
  asigs.forEach(a => {
    const opt = document.createElement('option');
    opt.value = opt.textContent = a;
    selAsig.appendChild(opt);
  });

  document.getElementById('obs-field-estudiante').style.display = 'flex';
  cargarEstudiantesEnLista(grado);
});

document.getElementById('obs-select-asignatura').addEventListener('change', verificarFormCompleto);

// ── Búsqueda de estudiantes ───────────────────────────────
function cargarEstudiantesEnLista(grado) {
  obsEstado.estudianteFiltrado = (obsEstado.estudiantesPorGrado[grado] || [])
    .slice()
    .sort((a, b) => a.nombre.localeCompare(b.nombre));
}

function mostrarListaEstudiantes(q) {
  const grado = document.getElementById('obs-select-grado').value;
  const todos = obsEstado.estudiantesPorGrado[grado] || [];
  const lista = document.getElementById('obs-lista-est');

  const filtrados = q
    ? todos.filter(e =>
        e.nombre.toLowerCase().includes(q) ||
        e.id.includes(q) ||
        (e.docId && e.docId.includes(q))
      )
    : [...todos].sort((a, b) => a.nombre.localeCompare(b.nombre));

  if (!filtrados.length) {
    lista.innerHTML = '<div class="obs-est-item"><span class="obs-est-item-nombre" style="color:#aaa;">Sin resultados</span></div>';
    lista.style.display = 'block';
    return;
  }

  lista.innerHTML = '';
  filtrados.slice(0, 30).forEach(e => {
    const div = document.createElement('div');
    div.className = 'obs-est-item';
    div.innerHTML = `<span class="obs-est-item-nombre">${e.nombre}</span><span class="obs-est-item-id">${e.docId ? 'Doc: ' + e.docId : ''}</span>`;
    div.addEventListener('click', () => seleccionarEstudiante(e));
    lista.appendChild(div);
  });
  lista.style.display = 'block';
}

document.getElementById('obs-buscar-est').addEventListener('focus', function() {
  mostrarListaEstudiantes(this.value.trim().toLowerCase());
});

document.getElementById('obs-buscar-est').addEventListener('input', function() {
  mostrarListaEstudiantes(this.value.trim().toLowerCase());
});

function seleccionarEstudiante(est) {
  obsEstado.estSeleccionado = est;
  document.getElementById('obs-lista-est').style.display = 'none';
  document.getElementById('obs-buscar-est').value = '';
  document.getElementById('obs-est-nombre-chip').textContent = est.nombre;
  document.getElementById('obs-est-chip').style.display = 'flex';
  document.getElementById('obs-buscar-est').closest('.obs-estudiante-search').style.display = 'none';
  verificarFormCompleto();
}

document.getElementById('obs-chip-clear').addEventListener('click', resetEstudiante);

function resetEstudiante() {
  obsEstado.estSeleccionado = null;
  document.getElementById('obs-est-chip').style.display = 'none';
  document.getElementById('obs-buscar-est').closest('.obs-estudiante-search').style.display = 'flex';
  document.getElementById('obs-buscar-est').value = '';
  document.getElementById('obs-lista-est').style.display = 'none';
  verificarFormCompleto();
}

// Cerrar lista al hacer clic fuera
document.addEventListener('click', function(e) {
  const lista  = document.getElementById('obs-lista-est');
  const search = document.getElementById('obs-buscar-est');
  if (lista && search && !lista.contains(e.target) && e.target !== search) {
    lista.style.display = 'none';
  }
});

// ── Verificar si el formulario está completo ──────────────
function verificarFormCompleto() {
  const periodo = document.getElementById('obs-select-periodo').value;
  const grado   = document.getElementById('obs-select-grado').value;
  const asig    = document.getElementById('obs-select-asignatura').value;
  const est     = obsEstado.estSeleccionado;

  const textoWrap  = document.getElementById('obs-field-texto');
  const actionWrap = document.getElementById('obs-action-wrap');
  const preview    = document.getElementById('obs-preview-info');

  if (periodo && grado && asig && est) {
    textoWrap.style.display  = 'flex';
    actionWrap.style.display = 'block';
    preview.innerHTML = `<strong>${est.nombre}</strong> · ${grado} · ${asig} · Periodo ${periodo}`;
  } else {
    textoWrap.style.display  = 'none';
    actionWrap.style.display = 'none';
  }
}

document.getElementById('obs-select-periodo').addEventListener('change', verificarFormCompleto);

// Contador de caracteres
document.getElementById('obs-textarea').addEventListener('input', function() {
  document.getElementById('obs-char-count').textContent = `${this.value.length} / 500`;
});

// ── Selector de tipo (Observación / Felicitación) ──
let obsTipoActual = 'Observación';

document.querySelectorAll('.obs-tipo-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.obs-tipo-btn').forEach(b => b.classList.remove('obs-tipo-activo'));
    btn.classList.add('obs-tipo-activo');
    obsTipoActual = btn.dataset.tipo;

    const esFelicitacion = obsTipoActual === 'Felicitación';
    document.getElementById('obs-label-texto').textContent   = obsTipoActual;
    document.getElementById('obs-btn-texto').textContent     = esFelicitacion ? 'Guardar felicitación' : 'Guardar observación';
    document.getElementById('obs-textarea').placeholder      = esFelicitacion
      ? 'Escribe aquí el mensaje de felicitación para el estudiante...'
      : 'Escriba aquí la observación sobre el estudiante...';
    document.getElementById('obs-btn-guardar').className = esFelicitacion
      ? 'obs-btn-guardar obs-btn-felicitacion'
      : 'obs-btn-guardar obs-btn-observacion';
  });
});

// ── Guardar observación ───────────────────────────────────
document.getElementById('obs-btn-guardar').addEventListener('click', guardarObservacion);

async function guardarObservacion() {
  const periodo = document.getElementById('obs-select-periodo').value;
  const grado   = document.getElementById('obs-select-grado').value;
  const asig    = document.getElementById('obs-select-asignatura').value;
  const est     = obsEstado.estSeleccionado;
  const texto   = document.getElementById('obs-textarea').value.trim();

  if (!periodo || !grado || !asig || !est || !texto) {
    mostrarObsFeedback('Por favor completa todos los campos antes de guardar.', 'err');
    return;
  }

  const hoy = new Date();
  const fecha = `${String(hoy.getDate()).padStart(2,'0')}/${String(hoy.getMonth()+1).padStart(2,'0')}/${hoy.getFullYear()}`;

  const payload = {
    tipo:          'guardarObservacion',
    periodo:       periodo,
    id_estudiante: est.docId || est.id,
    nombres:       est.nombre,
    observacion:   texto,
    fecha:         fecha,
    docente:       nombreDocente,
    grado:         grado,
    asignatura:    asig,
    tipoRegistro:  obsTipoActual
  };

  const btn = document.getElementById('obs-btn-guardar');
  btn.disabled = true;
  btn.style.opacity = '0.7';
  const textoOriginal = document.getElementById('obs-btn-texto')?.textContent || 'Guardar';
  if (document.getElementById('obs-btn-texto')) {
    document.getElementById('obs-btn-texto').textContent = 'Guardando...';
  }

  try {
    const resp = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      body:   JSON.stringify(payload)
    });
    const json = await resp.json();

    if (json.status === 'ok') {
      mostrarObsFeedback(`✓ ${obsTipoActual} guardada correctamente.`, 'ok');
      obsTipoActual = 'Observación';
      document.querySelectorAll('.obs-tipo-btn').forEach(b => b.classList.remove('obs-tipo-activo'));
      document.getElementById('obs-tipo-obs').classList.add('obs-tipo-activo');
      document.getElementById('obs-label-texto').textContent = 'Observación';
      document.getElementById('obs-btn-texto').textContent   = 'Guardar observación';
      document.getElementById('obs-textarea').placeholder    = 'Escriba aquí la observación sobre el estudiante...';
      document.getElementById('obs-btn-guardar').className   = 'obs-btn-guardar';

      // Limpiar formulario
      document.getElementById('obs-textarea').value = '';
      document.getElementById('obs-char-count').textContent = '0 / 500';
      resetEstudiante();
      document.getElementById('obs-select-periodo').value = '';
      document.getElementById('obs-select-grado').value = '';
      document.getElementById('obs-select-asignatura').innerHTML = '<option value="">— elige grado primero —</option>';
      document.getElementById('obs-field-estudiante').style.display = 'none';
      document.getElementById('obs-field-texto').style.display = 'none';
      document.getElementById('obs-action-wrap').style.display = 'none';
      // Invalidar caché de consulta
      obsEstado.todasObs = [];
    } else {
      mostrarObsFeedback('Error al guardar: ' + (json.mensaje || 'intente de nuevo.'), 'err');
    }
  } catch(err) {
    mostrarObsFeedback('Error de conexión. Verifique su red e intente de nuevo.', 'err');
  } finally {
    btn.disabled = false;
    btn.style.opacity = '1';
    if (document.getElementById('obs-btn-texto')) {
      document.getElementById('obs-btn-texto').textContent = textoOriginal;
    }
  }
}

function mostrarObsFeedback(msg, tipo) {
  const el = document.getElementById('obs-feedback');
  el.textContent = msg;
  el.className   = 'obs-feedback ' + tipo;
  el.style.display = 'flex';
  setTimeout(() => { el.style.display = 'none'; }, 5000);
}

// ── Consultar observaciones ───────────────────────────────
async function cargarObservacionesConsulta() {
  const loading = document.getElementById('obs-consulta-loading');
  const lista   = document.getElementById('obs-lista-registros');
  const vacia   = document.getElementById('obs-consulta-vacia');

  loading.style.display = 'flex';
  lista.innerHTML = '';
  vacia.style.display = 'none';

  try {
    if (!obsEstado.todasObs.length) {
      const data = await fetchSheet(SHEET_OBS);
      if (data.error || !data.values || data.values.length < 2) {
        loading.style.display = 'none';
        vacia.style.display = 'flex';
        return;
      }

      const headers = data.values[0].map(h => (h || '').trim());
      const rows    = data.values.slice(1);

      const iP   = headers.findIndex(h => norm(h).includes('PERIODO'));
      const iId  = headers.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
      const iNom = headers.findIndex(h => norm(h).includes('NOMBRE'));
      const iObs = headers.findIndex(h => norm(h).includes('OBSERVA'));
      const iFec = headers.findIndex(h => norm(h).includes('FECHA'));
      const iDoc = headers.findIndex(h => norm(h).includes('DOCENTE'));
      const iGra = headers.findIndex(h => norm(h) === 'GRADO');

      // Grados donde el docente es director (tiene su nombre en la columna comportamiento)
      const gradosComoDirector = Object.keys(obsEstado.estudiantesPorGrado).filter(g => {
        const headers2 = data.values[0].map(h => (h || '').trim());
        const rows2    = data.values.slice(1);
        const iComp    = headers2.findIndex(h => h.toLowerCase().trim() === 'comportamiento');
        const iGrado2  = headers2.findIndex(h => h.toUpperCase() === 'GRADO');
        if (iComp < 0) return false;
        return rows2.some(row =>
          (row[iGrado2] || '').trim() === g &&
          (row[iComp]   || '').trim() === nombreDocente.trim()
        );
      });

      obsEstado.todasObs = rows
        .filter(r => {
          const autor = (r[iDoc] || '').trim();
          const grado = iGra >= 0 ? (r[iGra] || '').trim() : '';
          // Ver si es el autor, o si es director del grado de esa observación
          return autor === nombreDocente.trim() || gradosComoDirector.includes(grado);
        })
        .map(r => ({
          periodo: (r[iP]   || '').trim(),
          id:      (r[iId]  || '').trim(),
          nombre:  (r[iNom] || '').trim(),
          obs:     (r[iObs] || '').trim(),
          fecha:   (r[iFec] || '').trim(),
          docente: (r[iDoc] || '').trim(),
          grado:   iGra >= 0 ? (r[iGra] || '').trim() : '',
          tipoRegistro: (() => { const iTipo = headers.findIndex(h => norm(h) === 'TIPO'); return iTipo >= 0 ? (r[iTipo] || 'Observación').trim() : 'Observación'; })()
        }))
        .filter(r => r.id || r.nombre);
    }

    loading.style.display = 'none';
    aplicarFiltrosConsulta();

  } catch(err) {
    loading.style.display = 'none';
    lista.innerHTML = '<div style="padding:20px;text-align:center;color:#aaa;font-size:13px;">Error de conexión al cargar observaciones.</div>';
  }
}

function aplicarFiltrosConsulta() {
  const filtGrado   = document.getElementById('obs-filtro-grado').value;
  const filtPeriodo = document.getElementById('obs-filtro-periodo').value;
  const filtNombre  = document.getElementById('obs-filtro-nombre').value.trim().toLowerCase();
  const lista       = document.getElementById('obs-lista-registros');
  const vacia       = document.getElementById('obs-consulta-vacia');

  const filtradas = obsEstado.todasObs.filter(r => {
    const okGrado   = !filtGrado   || r.grado   === filtGrado;
    const okPeriodo = !filtPeriodo || r.periodo  === filtPeriodo;
    const okNombre  = !filtNombre  || r.nombre.toLowerCase().includes(filtNombre) || r.id.includes(filtNombre);
    return okGrado && okPeriodo && okNombre;
  });

  lista.innerHTML = '';

  if (!filtradas.length) {
    vacia.style.display = 'flex';
    return;
  }

  vacia.style.display = 'none';

  filtradas.slice().reverse().forEach(r => {
    const card = document.createElement('div');
    const esFel = r.tipoRegistro === 'Felicitación';
    card.className = 'obs-registro-card' + (esFel ? ' obs-registro-felicitacion' : '');
    card.innerHTML = `
      <div class="obs-registro-top">
        <div>
          <div class="obs-registro-nombre">${r.nombre}</div>
          <div class="obs-registro-id">Doc. ${r.id}${r.grado ? ' · ' + r.grado : ''}</div>
        </div>
        <div class="obs-registro-badges">
          ${r.grado ? `<span class="obs-badge obs-badge-grado">${r.grado}</span>` : ''}
          ${r.periodo ? `<span class="obs-badge obs-badge-periodo">Periodo ${r.periodo}</span>` : ''}
          <span class="obs-badge ${esFel ? 'obs-badge-felicitacion' : 'obs-badge-obs'}">${r.tipoRegistro || 'Observación'}</span>
        </div>
      </div>
      <div class="obs-registro-texto">${r.obs}</div>
      <div class="obs-registro-footer">
        <span>
          <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
          </svg>
          ${r.fecha}
        </span>
        <span>
          <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
          </svg>
          ${r.docente}
        </span>
      </div>
    `;
    lista.appendChild(card);
  });
}

// Listeners de filtros de consulta
['obs-filtro-grado','obs-filtro-periodo'].forEach(id => {
  document.getElementById(id).addEventListener('change', aplicarFiltrosConsulta);
});
document.getElementById('obs-filtro-nombre').addEventListener('input', aplicarFiltrosConsulta);