// ════════════════════════════════════════════════════════
//  CONFIGURACIÓN
// ════════════════════════════════════════════════════════
const API_KEY           = "AIzaSyA1GsdvskVoIQDSOg6EI8aEQW5MnQsFeoE";
const SPREADSHEET_ID    = "1ytP_bXfbjRsA05KJYM8HAxQvTvnAeHgz09r4QhMy-Rs";
const SHEET_ESTUDIANTES = "asignaturas_estudiantes";
const SHEET_PERIODOS    = "Periodos";
const SHEET_INDICADORES = "Indicadores";

const APPS_SCRIPT_URL   = "https://script.google.com/macros/s/AKfycbygBydUcKQWczkcs4jL3BshWNY1npLtsqK9IbNA72nIgpRhmjHo3tKnCVry0AcwvGIf/exec";

// Columnas de asignaturas_estudiantes que NO son asignaturas
const METADATOS = [
  'no_identificador','nombres_apellidos','grado',
  'sede','estado','comportamiento'
];

const ORDEN_GRADOS = [
  'Preescolar','Primero','Segundo','Tercero','Cuarto','Quinto',
  'Sexto','Séptimo','Octavo','Noveno','Décimo','Undécimo'
];

// ════════════════════════════════════════════════════════
//  ESTADO GLOBAL
// ════════════════════════════════════════════════════════
let nombreDocente        = "";
let sedeDocente          = "";
let mapaGradosAsigs      = {};   // { grado: [asig1, asig2, ...] }
let periodos             = [];   // [{ num, inicio, fin, cierreSistema }]
let indicadoresGuardados = {};   // { "PER|GRADO|ASIG": "texto en Sheets" }
let bufferCambios        = {};   // { "PER|GRADO|ASIG": "texto en sesión" }
let periodoActivo        = null;
let gradoActivo          = null;
let periodoBloqueado     = false;

// ════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════
function getUrlParameter(name) {
  name = name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]');
  const regex   = new RegExp('[\\?&]' + name + '=([^&#]*)');
  const results = regex.exec(location.search);
  return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
}

function fetchSheet(sheetName) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(sheetName)}?key=${API_KEY}`;
  return fetch(url).then(r => r.json());
}

function norm(h) {
  return String(h || '').trim().toUpperCase()
    .replace(/[ÁÀÂÄ]/g,'A').replace(/[ÉÈÊË]/g,'E')
    .replace(/[ÍÌÎÏ]/g,'I').replace(/[ÓÒÔÖ]/g,'O').replace(/[ÚÙÛÜ]/g,'U')
    .replace(/[áàâä]/g,'A').replace(/[éèêë]/g,'E')
    .replace(/[íìîï]/g,'I').replace(/[óòôö]/g,'O').replace(/[úùûü]/g,'U')
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

function claveIndicador(periodo, grado, asignatura) {
  return `${periodo}|${grado}|${asignatura}`;
}

function iniciales(nombre) {
  return nombre.split(' ')
    .filter(p => p.length > 2)
    .slice(0, 2)
    .map(p => p[0].toUpperCase())
    .join('') || nombre.substring(0, 2).toUpperCase();
}

function setSaveStatus(msg, tipo) {
  const el = document.getElementById('save-status');
  if (!el) return;
  el.textContent = msg;
  el.className = 'save-status' + (tipo ? ' ' + tipo : '');
}

function mostrarSaveMsg(msg, tipo) {
  const el = document.getElementById('save-msg');
  if (!el) return;
  el.textContent = msg;
  el.className   = `save-msg ${tipo}`;
  el.style.display = 'flex';
  setTimeout(() => { el.style.display = 'none'; }, 4000);
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
    window.location.href = 'Index.html';
  }
});

// ════════════════════════════════════════════════════════
//  INICIALIZACIÓN PRINCIPAL
// ════════════════════════════════════════════════════════
async function inicializar() {
  try {
    const [dataEst, dataPer, dataInd] = await Promise.all([
      fetchSheet(SHEET_ESTUDIANTES),
      fetchSheet(SHEET_PERIODOS),
      fetchSheet(SHEET_INDICADORES)
    ]);

    procesarPeriodos(dataPer);
    procesarEstudiantes(dataEst);
    procesarIndicadoresGuardados(dataInd);

    const grados = Object.keys(mapaGradosAsigs);
    if (!grados.length) {
      document.getElementById('state-loading-init').style.display = 'none';
      document.getElementById('state-sin-asignaturas').style.display = 'flex';
      return;
    }

    document.getElementById('state-loading-init').style.display = 'none';
    document.getElementById('panel-principal').style.display     = 'block';

    construirPestañasGrados(grados);
    vincularTabsPeriodo();

    const periodoDefecto = detectarPeriodoActivo();
    activarPeriodo(periodoDefecto);

  } catch (err) {
    console.error('Error en inicialización:', err);
    document.getElementById('state-loading-init').innerHTML =
      '<p style="color:#c62828;">Error al cargar los datos. Verifica tu conexión.</p>';
  }
}

// ════════════════════════════════════════════════════════
//  PROCESAR PERIODOS
//  El bloqueo usa ÚNICAMENTE cierreSistema (col "CIERRE SISTEMA").
//  El fin académico del periodo NO bloquea la edición.
// ════════════════════════════════════════════════════════
function procesarPeriodos(dataPer) {
  if (dataPer.error || !dataPer.values || dataPer.values.length < 2) return;

  const headers = dataPer.values[0].map(h => (h || '').trim());

  const iNum = headers.findIndex(h => norm(h).includes('PERIODO'));

  const iIni = headers.findIndex(h => norm(h).includes('INICIO'));

  // Fin académico: contiene FIN o FINAL pero NO contiene CIERRE ni SISTEMA
  const iFin = headers.findIndex(h =>
    (norm(h).includes('FINAL') || norm(h).includes('FIN')) &&
    !norm(h).includes('CIERRE') && !norm(h).includes('SISTEMA')
  );

  // Cierre del sistema: debe contener CIERRE y SISTEMA
  const iCieSistema = headers.findIndex(h =>
    norm(h).includes('CIERRE') && norm(h).includes('SISTEMA')
  );

  periodos = dataPer.values.slice(1).map(row => ({
    num:           String(row[iNum]          || '').trim(),
    inicio:        iIni        >= 0 ? parseFecha(row[iIni])        : null,
    fin:           iFin        >= 0 ? parseFecha(row[iFin])         : null,
    cierreSistema: iCieSistema >= 0 ? parseFecha(row[iCieSistema])  : null
  })).filter(p => p.num);
}

// ════════════════════════════════════════════════════════
//  PROCESAR ESTUDIANTES
// ════════════════════════════════════════════════════════
function procesarEstudiantes(dataEst) {
  if (dataEst.error || !dataEst.values || dataEst.values.length < 2) return;

  const headers = dataEst.values[0].map(h => (h || '').trim());
  const rows    = dataEst.values.slice(1);

  const iGrado  = headers.findIndex(h => h.toUpperCase() === 'GRADO');
  const iEstado = headers.findIndex(h => h.toUpperCase() === 'ESTADO');

  const colsAsig = headers
    .map((h, i) => ({ nombre: h, idx: i }))
    .filter(c => !METADATOS.includes(c.nombre.toLowerCase().trim()));

  const filasDocente = rows.filter(row => {
    const estado = (row[iEstado] || '').trim().toLowerCase();
    if (estado === 'inactivo' || estado === 'retirado') return false;
    return colsAsig.some(c => (row[c.idx] || '').trim() === nombreDocente.trim());
  });

  mapaGradosAsigs = {};
  filasDocente.forEach(row => {
    const grado = (row[iGrado] || '').trim();
    if (!grado) return;
    colsAsig.forEach(c => {
      if ((row[c.idx] || '').trim() === nombreDocente.trim()) {
        if (!mapaGradosAsigs[grado]) mapaGradosAsigs[grado] = new Set();
        mapaGradosAsigs[grado].add(c.nombre);
      }
    });
  });

  // Comportamiento solo se agrega si el docente figura en esa columna para ese grado
  const iComportamiento = headers.findIndex(h => h.toLowerCase().trim() === 'comportamiento');
  Object.keys(mapaGradosAsigs).forEach(g => {
    const arr = [...mapaGradosAsigs[g]];
    if (iComportamiento >= 0) {
      const esDirGrado = rows.some(row =>
        (row[iGrado] || '').trim() === g &&
        (row[iComportamiento] || '').trim() === nombreDocente.trim()
      );
      if (esDirGrado) arr.push('Comportamiento');
    }
    mapaGradosAsigs[g] = arr;
  });
}

// ════════════════════════════════════════════════════════
//  PROCESAR INDICADORES GUARDADOS
// ════════════════════════════════════════════════════════
function procesarIndicadoresGuardados(dataInd) {
  indicadoresGuardados = {};
  if (dataInd.error || !dataInd.values || dataInd.values.length < 2) return;

  const headers = dataInd.values[0].map(h => (h || '').trim());
  const rows    = dataInd.values.slice(1);

  const iPer  = headers.findIndex(h => norm(h) === 'PERIODO');
  const iGra  = headers.findIndex(h => norm(h) === 'GRADO');
  const iAsig = headers.findIndex(h => norm(h) === 'ASIGNATURA');
  const iDoc  = headers.findIndex(h => norm(h) === 'DOCENTE');
  const iInd  = headers.findIndex(h => norm(h) === 'INDICADOR');

  rows.forEach(row => {
    if ((row[iDoc] || '').trim() !== nombreDocente.trim()) return;
    const per  = String(row[iPer]  || '').trim();
    const gra  = String(row[iGra]  || '').trim();
    const asig = String(row[iAsig] || '').trim();
    const ind  = String(row[iInd]  || '').trim();
    if (per && gra && asig) {
      indicadoresGuardados[claveIndicador(per, gra, asig)] = ind;
    }
  });
}

// ════════════════════════════════════════════════════════
//  DETECTAR PERIODO ACTIVO POR DEFECTO
// ════════════════════════════════════════════════════════
function detectarPeriodoActivo() {
  const hoy = new Date(); hoy.setHours(0,0,0,0);

  // Periodo cuyo sistema está abierto ahora
  for (const p of periodos) {
    if (!p.inicio) continue;
    const cierreEfectivo = p.cierreSistema || p.fin;
    if (hoy >= p.inicio && (!cierreEfectivo || hoy <= cierreEfectivo)) return p.num;
  }

  // Si ninguno abierto, el más próximo al futuro
  const futuros = periodos.filter(p => p.inicio && p.inicio > hoy);
  if (futuros.length) return futuros[0].num;

  return periodos.length ? periodos[periodos.length - 1].num : '1';
}

// ════════════════════════════════════════════════════════
//  CONSTRUIR PESTAÑAS DE GRADOS
// ════════════════════════════════════════════════════════
function construirPestañasGrados(grados) {
  const gradosOrdenados = ordenarGrados(grados);
  const contenedor = document.getElementById('grados-tabs');
  contenedor.innerHTML = '';

  gradosOrdenados.forEach(grado => {
    const btn = document.createElement('button');
    btn.className     = 'grado-tab';
    btn.dataset.grado = grado;

    const span = document.createElement('span');
    span.textContent = grado;

    const badge = document.createElement('span');
    badge.className   = 'tab-badge';
    badge.id          = `badge-${grado.replace(/\s/g,'_')}`;
    badge.textContent = '—';

    btn.appendChild(span);
    btn.appendChild(badge);
    btn.addEventListener('click', () => activarGrado(grado));
    contenedor.appendChild(btn);
  });

  actualizarBadgesGrados();
}

// ════════════════════════════════════════════════════════
//  ACTUALIZAR BADGES (buffer tiene prioridad sobre Sheets)
// ════════════════════════════════════════════════════════
function actualizarBadgesGrados() {
  if (!periodoActivo) return;

  Object.keys(mapaGradosAsigs).forEach(grado => {
    const asigs = mapaGradosAsigs[grado] || [];
    const total = asigs.length;

    const conTexto = asigs.filter(a => {
      const clave = claveIndicador(periodoActivo, grado, a);
      const texto = bufferCambios.hasOwnProperty(clave)
        ? bufferCambios[clave]
        : (indicadoresGuardados[clave] || '');
      return texto.trim() !== '';
    }).length;

    const badge = document.getElementById(`badge-${grado.replace(/\s/g,'_')}`);
    if (!badge) return;
    badge.textContent = total > 0 ? `${conTexto}/${total}` : '—';
    badge.classList.toggle('completo', conTexto === total && total > 0);
  });
}

// ════════════════════════════════════════════════════════
//  VINCULAR TABS DE PERIODO
// ════════════════════════════════════════════════════════
function vincularTabsPeriodo() {
  document.querySelectorAll('.periodo-tab').forEach(btn => {
    btn.addEventListener('click', () => activarPeriodo(btn.dataset.periodo));
  });
}

// ════════════════════════════════════════════════════════
//  ACTIVAR PERIODO
//  Bloquea solo si: hoy < inicio  o  hoy > cierreSistema
//  El fin académico del periodo NO bloquea.
// ════════════════════════════════════════════════════════
function activarPeriodo(numPeriodo) {
  periodoActivo = numPeriodo;

  document.querySelectorAll('.periodo-tab').forEach(btn => {
    btn.classList.toggle('activo', btn.dataset.periodo === numPeriodo);
  });

  const hoy       = new Date(); hoy.setHours(0,0,0,0);
  const per       = periodos.find(p => p.num === numPeriodo);
  const badgeEl   = document.getElementById('badge-estado-texto');
  const bannerEl  = document.getElementById('banner-bloqueado');
  const bannerMsg = document.getElementById('banner-bloqueado-msg');

  document.getElementById('periodo-estado').style.display = 'flex';
  bannerEl.style.display = 'none';
  periodoBloqueado = false;

  if (!per || !per.inicio) {
    // Sin fechas → editable, sin restricción
    badgeEl.textContent = per ? 'Sin fechas configuradas' : 'Periodo ' + numPeriodo;
    badgeEl.className   = 'badge-estado futuro';

  } else if (hoy < per.inicio) {
    // Antes de inicio → bloqueado
    const dias = Math.ceil((per.inicio - hoy) / 86400000);
    badgeEl.textContent = `Inicia en ${dias} día${dias !== 1 ? 's' : ''}`;
    badgeEl.className   = 'badge-estado futuro';
    periodoBloqueado    = true;
    bannerEl.style.display = 'flex';
    bannerMsg.textContent  = 'Este periodo aún no ha iniciado — los indicadores no se pueden editar.';

  } else if (per.cierreSistema && hoy > per.cierreSistema) {
    // Pasó el cierre del sistema → bloqueado
    badgeEl.textContent = 'Sistema cerrado';
    badgeEl.className   = 'badge-estado cerrado';
    periodoBloqueado    = true;
    bannerEl.style.display = 'flex';
    bannerMsg.textContent  = 'El sistema de ingreso está cerrado para este periodo — solo lectura.';

  } else {
    // Sistema abierto → editable
    if (per.cierreSistema) {
      const dias = Math.ceil((per.cierreSistema - hoy) / 86400000);
      badgeEl.textContent = `Abierto · cierra en ${dias} día${dias !== 1 ? 's' : ''}`;
    } else if (per.fin && hoy > per.fin) {
      // Periodo terminado pero sin cierre sistema → seguimos abiertos
      badgeEl.textContent = 'Periodo finalizado · sistema abierto';
    } else {
      badgeEl.textContent = 'En curso · Abierto';
    }
    badgeEl.className = 'badge-estado abierto';
  }

  actualizarBadgesGrados();

  if (gradoActivo) {
    renderTablaIndicadores();
  } else {
    const primerGrado = ordenarGrados(Object.keys(mapaGradosAsigs))[0];
    if (primerGrado) activarGrado(primerGrado);
  }
}

// ════════════════════════════════════════════════════════
//  ACTIVAR GRADO
// ════════════════════════════════════════════════════════
function activarGrado(grado) {
  gradoActivo = grado;
  document.querySelectorAll('.grado-tab').forEach(btn => {
    btn.classList.toggle('activo', btn.dataset.grado === grado);
  });
  renderTablaIndicadores();
}

// ════════════════════════════════════════════════════════
//  RENDERIZAR TABLA
//  Valor mostrado: bufferCambios > indicadoresGuardados > vacío
// ════════════════════════════════════════════════════════
function renderTablaIndicadores() {
  if (!gradoActivo || !periodoActivo) return;

  const btnGuardar = document.getElementById('btn-guardar');
  const saveBar    = document.getElementById('save-bar');

  document.getElementById('state-loading-indicadores').style.display = 'none';
  document.getElementById('indicadores-card').style.display = 'block';

  document.getElementById('card-titulo').textContent =
    `${gradoActivo} — Periodo ${periodoActivo}`;
  document.getElementById('card-subtitulo').textContent =
    'Indicadores de logro por asignatura';

  btnGuardar.style.display = periodoBloqueado ? 'none' : 'flex';
  saveBar.style.display    = periodoBloqueado ? 'none' : 'block';
  document.getElementById('instruccion').style.display = periodoBloqueado ? 'none' : 'flex';

  const asignaturas = mapaGradosAsigs[gradoActivo] || [];
  const tbody       = document.getElementById('tbody-indicadores');
  tbody.innerHTML   = '';

  asignaturas.forEach((asig, idx) => {
    const clave = claveIndicador(periodoActivo, gradoActivo, asig);

    // Prioridad: buffer de sesión > Sheets > vacío
    const enBuffer     = bufferCambios.hasOwnProperty(clave) ? bufferCambios[clave] : null;
    const enSheets     = indicadoresGuardados[clave] || '';
    const valorMostrar = enBuffer !== null ? enBuffer : enSheets;
    const tieneValor   = valorMostrar.trim() !== '';

    const tr = document.createElement('tr');

    // Asignatura
    const tdAsig = document.createElement('td');
    tdAsig.style.paddingTop = '14px';
    tdAsig.innerHTML = `
      <div class="td-asig-nombre">
        <div class="asig-icono">${iniciales(asig)}</div>
        <span class="asig-nombre">${asig}</span>
      </div>`;

    // Textarea
    const tdInd = document.createElement('td');
    tdInd.className = 'td-indicador';

    const textarea = document.createElement('textarea');
    textarea.className   = 'indicador-textarea';
    textarea.placeholder = periodoBloqueado
      ? 'No hay indicador registrado para este periodo.'
      : 'Escribe aquí el indicador de logro...';
    textarea.value        = valorMostrar;
    textarea.maxLength    = 500;
    textarea.disabled     = periodoBloqueado;
    textarea.dataset.asig     = asig;
    textarea.dataset.clave    = clave;
    textarea.dataset.original = enSheets;
    if (tieneValor) textarea.classList.add('tiene-datos');

    const counter = document.createElement('div');
    counter.className   = 'char-counter';
    counter.textContent = `${valorMostrar.length}/500`;

    if (!periodoBloqueado) {
      textarea.addEventListener('input', () => {
        const len = textarea.value.length;
        counter.textContent = `${len}/500`;
        counter.classList.toggle('limite', len >= 480);
        textarea.classList.toggle('tiene-datos', textarea.value.trim() !== '');
        actualizarChipEstado(tr, textarea.value.trim(), textarea.dataset.original.trim());
        setSaveStatus('');
        // Guardar en buffer multi-grado
        bufferCambios[clave] = textarea.value;
        actualizarBadgesGrados();
      });
    }

    tdInd.appendChild(textarea);
    if (!periodoBloqueado) tdInd.appendChild(counter);

    // Chip de estado
    const tdEstado = document.createElement('td');
    tdEstado.className = 'td-estado-cell';

    const chip = document.createElement('span');
    chip.className = 'chip-estado';

    if (enBuffer !== null && enBuffer.trim() !== enSheets.trim()) {
      chip.classList.add('pendiente');
      chip.innerHTML = '✎ Sin guardar';
    } else if (tieneValor) {
      chip.classList.add('guardado');
      chip.innerHTML = '✓ Guardado';
    } else {
      chip.classList.add('vacio');
      chip.innerHTML = '— Vacío';
    }

    tdEstado.appendChild(chip);
    tr.appendChild(tdAsig);
    tr.appendChild(tdInd);
    tr.appendChild(tdEstado);
    tbody.appendChild(tr);
  });

  setSaveStatus('');
}

// ════════════════════════════════════════════════════════
//  CHIP DE ESTADO
// ════════════════════════════════════════════════════════
function actualizarChipEstado(tr, valorActual, valorOriginal) {
  const chip = tr.querySelector('.chip-estado');
  if (!chip) return;
  chip.className = 'chip-estado';
  if (!valorActual) {
    chip.classList.add('vacio');
    chip.innerHTML = '— Vacío';
  } else if (valorActual !== valorOriginal) {
    chip.classList.add('pendiente');
    chip.innerHTML = '✎ Sin guardar';
  } else {
    chip.classList.add('guardado');
    chip.innerHTML = '✓ Guardado';
  }
}

// ════════════════════════════════════════════════════════
//  RECOPILAR DESDE BUFFER (todos los grados del periodo)
// ════════════════════════════════════════════════════════
function recopilarIndicadores() {
  const fecha     = new Date().toLocaleDateString('es-CO');
  const registros = [];

  Object.entries(bufferCambios).forEach(([clave, texto]) => {
    const partes = clave.split('|');
    if (partes.length < 3) return;
    const [periodo, grado, asignatura] = partes;
    if (periodo !== periodoActivo) return;
    registros.push({
      periodo,
      grado,
      asignatura,
      docente:   nombreDocente,
      fecha,
      indicador: texto.trim()
    });
  });

  return registros;
}

// ════════════════════════════════════════════════════════
//  ENVIAR A APPS SCRIPT
// ════════════════════════════════════════════════════════
function enviarAAppsScript(payload) {
  return new Promise((resolve, reject) => {
    fetch(APPS_SCRIPT_URL, {
      method:   'POST',
      headers:  { 'Content-Type': 'text/plain' },
      body:     JSON.stringify(payload),
      redirect: 'follow'
    })
    .then(r => r.text().then(text => {
      try { resolve(JSON.parse(text)); }
      catch { resolve({ status: 'ok', mensaje: 'Datos enviados.' }); }
    }))
    .catch(() => enviarConIframe(JSON.stringify(payload)).then(resolve).catch(reject));
  });
}

function enviarConIframe(payloadStr) {
  return new Promise(resolve => {
    const iframeName = 'sianario_ind_' + Date.now();
    const iframe = document.createElement('iframe');
    iframe.name  = iframeName;
    iframe.style.display = 'none';
    document.body.appendChild(iframe);

    const form = document.createElement('form');
    form.method  = 'POST';
    form.action  = APPS_SCRIPT_URL;
    form.target  = iframeName;
    form.enctype = 'application/x-www-form-urlencoded';

    const inp = document.createElement('input');
    inp.type  = 'hidden';
    inp.name  = 'payload';
    inp.value = payloadStr;
    form.appendChild(inp);
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
      catch { resolve({ status: 'ok', mensaje: 'Datos enviados.' }); }
      limpiar();
    };
    form.submit();
  });
}

// ════════════════════════════════════════════════════════
//  GUARDAR INDICADORES
// ════════════════════════════════════════════════════════
async function guardarIndicadores() {
  if (periodoBloqueado) {
    setSaveStatus('🔒 No se puede guardar: el periodo está cerrado.', 'error');
    return;
  }

  const registros = recopilarIndicadores();
  if (!registros.length) {
    setSaveStatus('No hay cambios para guardar. Escribe al menos un indicador.', 'error');
    return;
  }

  const btnTop = document.getElementById('btn-guardar');
  const btnBot = document.getElementById('btn-guardar-bottom');
  btnTop.disabled = true;
  if (btnBot) btnBot.disabled = true;
  setSaveStatus('Guardando indicadores...', '');

  try {
    const result = await enviarAAppsScript({ tipo: 'guardarIndicadores', registros });

    if (result.status === 'ok') {
      // Sincronizar caché con Sheets
      registros.forEach(r => {
        indicadoresGuardados[claveIndicador(r.periodo, r.grado, r.asignatura)] = r.indicador;
      });

      // Limpiar buffer del periodo activo
      Object.keys(bufferCambios).forEach(clave => {
        if (clave.startsWith(periodoActivo + '|')) delete bufferCambios[clave];
      });

      setSaveStatus(
        `✓ ${result.guardados || registros.length} indicador(es) guardados correctamente.`,
        'success'
      );
      mostrarSaveMsg('✅ Indicadores guardados correctamente.', 'success');

      // Actualizar chips visibles
      document.querySelectorAll('.indicador-textarea').forEach(ta => {
        ta.dataset.original = ta.value.trim();
        actualizarChipEstado(ta.closest('tr'), ta.value.trim(), ta.value.trim());
        if (ta.value.trim()) ta.classList.add('tiene-datos');
      });

      actualizarBadgesGrados();

    } else {
      setSaveStatus(`Error: ${result.mensaje}`, 'error');
      mostrarSaveMsg(`⚠️ ${result.mensaje}`, 'error');
    }

  } catch (err) {
    setSaveStatus('Error al enviar datos. Verifica tu conexión.', 'error');
    mostrarSaveMsg('⚠️ Error de conexión. Intenta nuevamente.', 'error');
    console.error(err);
  }

  btnTop.disabled = false;
  if (btnBot) btnBot.disabled = false;
}

document.getElementById('btn-guardar').addEventListener('click', guardarIndicadores);
document.getElementById('btn-guardar-bottom').addEventListener('click', guardarIndicadores);

// ════════════════════════════════════════════════════════
//  ARRANCAR
// ════════════════════════════════════════════════════════
inicializar();