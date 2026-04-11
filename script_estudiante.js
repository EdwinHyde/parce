// ════════════════════════════════════════════════════════
//  PARCE — Portal Estudiante  v2.0
// ════════════════════════════════════════════════════════

const API_KEY        = "AIzaSyA1GsdvskVoIQDSOg6EI8aEQW5MnQsFeoE";
const SPREADSHEET_ID = "1ytP_bXfbjRsA05KJYM8HAxQvTvnAeHgz09r4QhMy-Rs";
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbygBydUcKQWczkcs4jL3BshWNY1npLtsqK9IbNA72nIgpRhmjHo3tKnCVry0AcwvGIf/exec";

function sheetUrl(sheetName) {
  return `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(sheetName)}?key=${API_KEY}`;
}

// ── Estado global ──
let G = {
  nombre:'', grado:'', sede:'', id:'',
  periodoActual:'1', periodoSelect:'1',
  consolidado:[], promGenerales:null,
  asistencia:[], observaciones:[], avisos:[], cuadroHonor:[],
};

// ════════════════════════════════════════════════════════
//  UTILIDADES
// ════════════════════════════════════════════════════════
function nivelNota(n) {
  const v = parseFloat(n);
  if (isNaN(v)) return 'sin';
  if (v >= 4.6) return 'medio';
  if (v >= 3.9) return 'alto';
  if (v >= 3.0) return 'basico';
  return 'bajo';
}

function emojiAsig(nombre) {
  const m = {
    'matem':'📐','lengua':'📖','castellan':'📖','ingles':'🌎',
    'ciencias nat':'🔬','ciencias soc':'🌍','fisica':'⚡',
    'quimica':'🧪','artistica':'🎨','arte':'🎨','etica':'🤝',
    'religion':'✝️','educacion fisica':'⚽','educ.fis':'⚽',
    'tecnolog':'💻','inform':'💻','filosofia':'🧠','estadist':'📊',
    'geometr':'📐','algebra':'🔢','emprendim':'💡','econom':'💰',
    'centro':'🎯','catedra':'🕊️','comportam':'🧑‍🤝‍🧑',
  };
  const low = (nombre||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  for (const [k,v] of Object.entries(m)) if (low.includes(k)) return v;
  return '📚';
}

function norm(h) {
  return String(h||'').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
function normH(h) { return norm(h); }

function getParam(name) {
  const r = new RegExp('[?&]' + name + '=([^&#]*)');
  const m = r.exec(location.search);
  return m ? decodeURIComponent(m[1].replace(/\+/g,' ')) : '';
}

function toast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2800);
}

function primerNombre(n) { return (n||'').split(' ')[0] || 'Estudiante'; }
function emptyHTML(icon, msg) { return `<div class="empty-state-card"><div class="empty-icon">${icon}</div><p>${msg}</p></div>`; }
function desempenoLabel(nota) {
  if (isNaN(nota)) return '';
  if (nota >= 4.6) return 'Superior';
  if (nota >= 4.0) return 'Alto';
  if (nota >= 3.0) return 'Básico';
  return 'Bajo';
}

function parseFechaLogin(str) {
  if (!str) return null;
  const dmy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) return new Date(+dmy[3], +dmy[2]-1, +dmy[1]);
  const ymd = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) return new Date(+ymd[1], +ymd[2]-1, +ymd[3]);
  if (/^\d+$/.test(str)) {
    const base = new Date(1899,11,30);
    base.setDate(base.getDate() + parseInt(str));
    return base;
  }
  return null;
}

function parseFechaRegistro(str) {
  if (!str) return null;
  const dmy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) return new Date(+dmy[3], +dmy[2]-1, +dmy[1]);
  const d = new Date(str);
  return isNaN(d) ? null : d;
}

function normalizarFechaLocal(raw) {
  const s = String(raw||'').trim();
  const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) return dmy[1].padStart(2,'0') + dmy[2].padStart(2,'0') + dmy[3];
  const ymd = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) return ymd[3] + ymd[2] + ymd[1];
  return s.replace(/\D/g,'');
}

// ════════════════════════════════════════════════════════
//  SISTEMA DE NOVEDADES
// ════════════════════════════════════════════════════════
const NOVEDAD_KEY = 'parce_visto_';

function marcarVisto(tab) {
  const key = NOVEDAD_KEY + tab + '_' + (G.id || G.nombre);
  localStorage.setItem(key, new Date().toISOString());
  quitarPunto(tab);
}

function quitarPunto(tab) {
  const pill = document.querySelector(`.pill[data-tab="${tab}"]`);
  if (!pill) return;
  const badge = pill.querySelector('.pill-badge');
  if (badge) badge.remove();
}

function mostrarPunto(tab) {
  const pill = document.querySelector(`.pill[data-tab="${tab}"]`);
  if (!pill || pill.querySelector('.pill-badge')) return;
  const dot = document.createElement('span');
  dot.className = 'pill-badge';
  pill.appendChild(dot);
}

function fechaUltimoVisto(tab) {
  const key = NOVEDAD_KEY + tab + '_' + (G.id || G.nombre);
  const val = localStorage.getItem(key);
  return val ? new Date(val) : null;
}

function verificarNovedades() {
  const periodo = G.periodoSelect;

  // Notas
  const vistoCal = fechaUltimoVisto('calificaciones');
  const filasNuevas = G.consolidado.filter(({ row, headers }) => {
    const iPer   = headers.findIndex(h => norm(h) === 'PERIODO');
    const iFecha = headers.findIndex(h => norm(h) === 'FECHA');
    if (iPer < 0 || String(row[iPer]||'').trim() !== periodo) return false;
    if (iFecha < 0) return !vistoCal;
    const f = parseFechaRegistro(String(row[iFecha]||''));
    return f && vistoCal && f > vistoCal;
  });
  if (filasNuevas.length > 0) mostrarPunto('calificaciones');

  // Anotaciones
  const vistoObs = fechaUltimoVisto('observaciones');
  const obsNuevas = G.observaciones.filter(({ row, headers, tipo }) => {
    if (tipo === 'Felicitación') return false;
    const iFecha = headers.findIndex(h => norm(h).includes('FECHA'));
    const f = iFecha >= 0 ? parseFechaRegistro(String(row[iFecha]||'')) : null;
    if (!f) return false;
    return !vistoObs || f > vistoObs;
  });
  if (obsNuevas.length > 0) mostrarPunto('observaciones');

  // Logros
  const vistoFel = fechaUltimoVisto('reconocimientos');
  const felNuevas = G.observaciones.filter(({ row, headers, tipo }) => {
    if (tipo !== 'Felicitación') return false;
    const iFecha = headers.findIndex(h => norm(h).includes('FECHA'));
    const f = iFecha >= 0 ? parseFechaRegistro(String(row[iFecha]||'')) : null;
    if (!f) return false;
    return !vistoFel || f > vistoFel;
  });
  if (felNuevas.length > 0) mostrarPunto('reconocimientos');

  // Asistencia: solo si ya visitó antes
  const vistoAsist = fechaUltimoVisto('asistencia');
  if (vistoAsist && G.asistencia.length > 0) {
    const { headers, row } = G.asistencia[0];
    const tieneFallas = Object.keys(NOMBRES_ASIG).some(codigo => {
      const idx = headers.findIndex(h => norm(h) === norm(codigo + '_P' + periodo));
      if (idx < 0) return false;
      const v = parseFloat(String(row[idx]||'0').replace(',','.'));
      return !isNaN(v) && v > 0;
    });
    if (tieneFallas) mostrarPunto('asistencia');
  }
}

// ════════════════════════════════════════════════════════
//  INIT
// ════════════════════════════════════════════════════════
window.addEventListener('DOMContentLoaded', () => {
  G.nombre = getParam('nombre') || 'Estudiante';
  G.grado  = getParam('grado')  || '';
  G.sede   = getParam('sede')   || '';

  const prog = document.getElementById('splash-progress');
  let pct = 0;
  const interval = setInterval(() => {
    pct = Math.min(pct + Math.random() * 18, 90);
    prog.style.width = pct + '%';
  }, 200);

  const h = new Date().getHours();
  const saludo = h < 12 ? '¡Buenos días!' : h < 18 ? '¡Buenas tardes!' : '¡Buenas noches!';
  document.getElementById('hero-greeting').textContent = saludo + ' 👋';
  document.getElementById('hero-name').textContent     = primerNombre(G.nombre);
  document.getElementById('hero-grado').textContent    = G.grado || 'Sin grado';
  document.getElementById('top-avatar').textContent    = (G.nombre[0]||'?').toUpperCase();

  setupTabs();
  setupPeriodoSelect();

  Promise.all([
    fetchConsolidado(),
    fetchAsistencia(),
    fetchObservaciones(),
    fetchAvisos(),
  ]).then(() => {
    clearInterval(interval);
    prog.style.width = '100%';
    setTimeout(() => {
      document.getElementById('splash').classList.add('fade-out');
      setTimeout(() => {
        document.getElementById('splash').style.display = 'none';
        document.getElementById('app').style.display    = 'flex';
        renderCalificaciones();
        renderAsistencia();
        renderReconocimientos();
        renderObservaciones();
        verificarNovedades();
        // Verificar novedades cada 3 minutos
        setInterval(async () => {
          G.observaciones = [];
          G.consolidado   = [];
          await fetchObservaciones();
          await fetchConsolidado();
          verificarNovedades();
        }, 3 * 60 * 1000);
      }, 500);
    }, 300);
  }).catch(err => {
    clearInterval(interval);
    prog.style.width = '100%';
    console.error(err);
    document.getElementById('splash').classList.add('fade-out');
    setTimeout(() => {
      document.getElementById('splash').style.display = 'none';
      document.getElementById('app').style.display    = 'flex';
      renderCalificaciones();
      renderAsistencia();
      renderReconocimientos();
      renderObservaciones();
    }, 500);
    toast('⚠️ Algunos datos no cargaron');
  });
});

// ════════════════════════════════════════════════════════
//  TABS Y PERIODO
// ════════════════════════════════════════════════════════
function setupTabs() {
  document.querySelectorAll('.pill').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
      marcarVisto(btn.dataset.tab);
      if (btn.dataset.tab === 'perfil') renderPerfil();
    });
  });
}

function setupPeriodoSelect() {
  const sel = document.getElementById('periodo-select');
  sel.addEventListener('change', () => {
    G.periodoSelect = sel.value;
    renderCalificaciones();
  });
}

// ════════════════════════════════════════════════════════
//  FETCH — CALIFICACIONES
// ════════════════════════════════════════════════════════
async function fetchConsolidado() {
  try {
    const resEst  = await fetch(sheetUrl('asignaturas_estudiantes'));
    const dataEst = await resEst.json();
    if (dataEst.values && dataEst.values.length > 1) {
      const hEst  = dataEst.values[0];
      const iId   = hEst.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
      const iNom  = hEst.findIndex(h => norm(h).includes('NOMBRE'));
      const iGrad = hEst.findIndex(h => norm(h) === 'GRADO');
      for (const row of dataEst.values.slice(1)) {
        const rowNom  = norm(String(row[iNom] ||''));
        const rowGrad = norm(String(row[iGrad]||''));
        const gradoN  = norm(G.grado), nomN = norm(G.nombre);
        const matchNom  = rowNom && nomN && (rowNom.includes(nomN.split(' ')[0]) || nomN.includes(rowNom.split(' ')[0]));
        const matchGrad = !gradoN || rowGrad.includes(gradoN) || gradoN.includes(rowGrad);
        if (matchNom && matchGrad) { G.id = iId >= 0 ? String(row[iId]||'').trim() : ''; break; }
      }
    }

    const res  = await fetch(sheetUrl('Calificaciones'));
    const data = await res.json();
    if (data.values && data.values.length > 1) {
      const headers  = data.values[0];
      const iIdCal   = headers.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
      const iNomCal  = headers.findIndex(h => norm(h) === 'NOMBRES Y APELLIDOS' || norm(h) === 'NOMBRES_APELLIDOS' || norm(h).includes('NOMBRE'));
      const iGradCal = headers.findIndex(h => norm(h) === 'GRADO');
      for (const row of data.values.slice(1)) {
        const rowId   = String(row[iIdCal]  ||'').trim();
        const rowNom  = norm(String(row[iNomCal] ||''));
        const rowGrad = norm(String(row[iGradCal]||''));
        const gradoN  = norm(G.grado);
        const matchId   = G.id && rowId === G.id;
        const matchNom  = rowNom && norm(G.nombre).split(' ')[0] && rowNom.includes(norm(G.nombre).split(' ')[0]);
        const matchGrad = !gradoN || rowGrad.includes(gradoN) || gradoN.includes(rowGrad);
        if ((matchId || matchNom) && matchGrad) G.consolidado.push({ headers, row });
      }
    }

    const resC  = await fetch(sheetUrl('Consolidado'));
    const dataC = await resC.json();
    if (dataC.values && dataC.values.length > 1) {
      const hC   = dataC.values[0];
      const iIdC = hC.findIndex(h => norm(h) === 'NO_IDENTIFICADOR');
      const iNomC= hC.findIndex(h => norm(h).includes('NOMBRE'));
      for (const row of dataC.values.slice(1)) {
        const rowId  = String(row[iIdC] ||'').trim();
        const rowNom = norm(String(row[iNomC]||''));
        const matchId  = G.id && rowId === G.id;
        const matchNom = rowNom && norm(G.nombre).split(' ')[0] && rowNom.includes(norm(G.nombre).split(' ')[0]);
        if (matchId || matchNom) { G.promGenerales = { headers: hC, row }; break; }
      }
    }

    await fetchPeriodoActivo();
  } catch(e) { console.error('Calificaciones:', e); }
}

async function fetchPeriodoActivo() {
  try {
    const res  = await fetch(sheetUrl('Periodos'));
    const data = await res.json();
    if (!data.values || data.values.length < 2) return;
    const headers = data.values[0].map(h => norm(h));
    const iNum    = headers.findIndex(h => h.includes('PERIODO') || h.includes('NUMERO'));
    const iInicio = headers.findIndex(h => h.includes('INICIO'));
    const iFinal  = headers.findIndex(h => h.includes('FINAL') || h.includes('FIN'));
    const hoy = new Date(); hoy.setHours(0,0,0,0);
    let encontrado = null;
    for (const row of data.values.slice(1)) {
      const num    = String(row[iNum]   ||'').trim();
      const inicio = parseFechaLogin(String(row[iInicio]||'').trim());
      const final  = parseFechaLogin(String(row[iFinal] ||'').trim());
      if (!num||!inicio||!final) continue;
      if (hoy >= inicio && hoy <= final) { encontrado = num; break; }
    }
    if (!encontrado) {
      for (const row of [...data.values.slice(1)].reverse()) {
        const num   = String(row[iNum] ||'').trim();
        const final = parseFechaLogin(String(row[iFinal]||'').trim());
        if (final && hoy > final) { encontrado = num; break; }
      }
    }
    if (!encontrado) encontrado = String(data.values[1][iNum]||'1').trim();
    G.periodoActual = encontrado;
    G.periodoSelect = encontrado;
    document.getElementById('periodo-select').value         = encontrado;
    document.getElementById('hero-periodo').textContent     = `Periodo ${encontrado}`;
  } catch(e) { console.error('Periodos:', e); }
}

// ════════════════════════════════════════════════════════
//  FETCH — ASISTENCIA
// ════════════════════════════════════════════════════════
async function fetchAsistencia() {
  try {
    const res  = await fetch(sheetUrl('CONSOLIDADO_ASISTENCIA'));
    const data = await res.json();
    if (!data.values || data.values.length < 2) return;
    const headers = data.values[0];
    const iId   = headers.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
    const iNoId = headers.findIndex(h => norm(h) === 'NO_IDENTIFICADOR');
    const iNom  = headers.findIndex(h => norm(h).includes('NOMBRE'));
    for (const row of data.values.slice(1)) {
      const rowId   = String(row[iId]  ||'').trim();
      const rowNoId = String(row[iNoId]||'').trim();
      const rowNom  = norm(String(row[iNom]||''));
      const matchId   = G.id && rowId === G.id;
      const matchNoId = G.id && rowNoId === G.id;
      const matchNom  = rowNom && norm(G.nombre).split(' ')[0] && rowNom.includes(norm(G.nombre).split(' ')[0]);
      if (matchId || matchNoId || matchNom) { G.asistencia = [{ headers, row }]; break; }
    }
  } catch(e) { console.error('Asistencia:', e); }
}

// ════════════════════════════════════════════════════════
//  FETCH — OBSERVACIONES
// ════════════════════════════════════════════════════════
async function fetchObservaciones() {
  try {
    const res  = await fetch(sheetUrl('Observaciones_Boletines'));
    const data = await res.json();
    if (!data.values || data.values.length < 2) return;
    const headers = data.values[0];
    const iId     = headers.findIndex(h => norm(h).includes('ID') || norm(h).includes('IDENTIFICADOR'));
    const iNombre = headers.findIndex(h => norm(h).includes('NOMBRE'));
    for (const row of data.values.slice(1)) {
      const rid     = String(row[iId]    ||'').trim();
      const rnombre = norm(String(row[iNombre]||''));
      const matchId     = G.id && rid === G.id.trim();
      const matchNombre = rnombre && norm(G.nombre).split(' ')[0] && rnombre.includes(norm(G.nombre).split(' ')[0]);
      if (matchId || matchNombre) {
        const iTipo = headers.findIndex(h => norm(h) === 'TIPO');
        const tipo  = iTipo >= 0 ? String(row[iTipo]||'Observación').trim() : 'Observación';
        G.observaciones.push({ headers, row, tipo });
      }
    }
  } catch(e) { console.error('Observaciones:', e); }
}

// ════════════════════════════════════════════════════════
//  FETCH — AVISOS
// ════════════════════════════════════════════════════════
async function fetchAvisos() {
  try {
    const res  = await fetch(sheetUrl('Avisos'));
    const data = await res.json();
    if (!data.values || data.values.length < 2) return;
    const headers  = data.values[0];
    const iDestino = headers.findIndex(h => norm(h).includes('DESTINATARIO') || norm(h).includes('DIRIGIDO'));
    const iActivo  = headers.findIndex(h => norm(h).includes('ACTIVO') || norm(h).includes('ESTADO'));
    const iTipo    = headers.findIndex(h => norm(h).includes('TIPO'));
    for (const row of data.values.slice(1)) {
      const destino    = norm(String(row[iDestino]||''));
      const activo     = norm(String(row[iActivo] ||''));
      const tipo       = norm(String(row[iTipo]   ||''));
      const visible    = destino==='' || destino.includes('TODO') || destino.includes('ESTUDIANTE');
      const estaActivo = activo==='' || activo==='TRUE' || activo==='ACTIVO' || activo==='1';
      if (visible && estaActivo) G.avisos.push({ headers, row, tipo });
    }
  } catch(e) { /* opcional */ }
}

// ════════════════════════════════════════════════════════
//  RENDER — CALIFICACIONES
// ════════════════════════════════════════════════════════
function renderCalificaciones() {
  const container = document.getElementById('calificaciones-list');
  const periodo   = G.periodoSelect;

  const filas = G.consolidado.filter(({ row, headers }) => {
    const iPer = headers.findIndex(h => norm(h) === 'PERIODO');
    return iPer >= 0 && String(row[iPer]||'').trim() === periodo;
  });

  if (filas.length === 0) {
    container.innerHTML = emptyHTML('📊', `No hay calificaciones registradas para el Periodo ${periodo}.`);
    document.getElementById('hero-promedio').textContent = '—';
    return;
  }

  if (G.promGenerales) {
    const { headers: hC, row: rC } = G.promGenerales;
    const iPromGen = hC.findIndex(h => norm(h) === norm(`PROMEDIO_GENERAL_P${G.periodoSelect}`));
    if (iPromGen >= 0) {
      const valProm = parseFloat(String(rC[iPromGen]||'').trim().replace(',','.'));
      if (!isNaN(valProm)) {
        const el = document.getElementById('hero-promedio');
        el.textContent = valProm.toFixed(1);
        el.className   = 'hero-nota-value ' + nivelNota(valProm);
      }
    }
  }

  let html = '';
  filas.forEach(({ headers, row }, i) => {
    const iAsig = headers.findIndex(h => norm(h) === 'ASIGNATURA');
    const iProm = headers.findIndex(h => norm(h).includes('PROMEDIO'));
    const iRec  = headers.findIndex(h => norm(h) === 'RECUPERACION');
    const iS1=headers.findIndex(h=>norm(h)==='SABER_1'||norm(h)==='SABER 1');
    const iS2=headers.findIndex(h=>norm(h)==='SABER_2'||norm(h)==='SABER 2');
    const iS3=headers.findIndex(h=>norm(h)==='SABER_3'||norm(h)==='SABER 3');
    const iS4=headers.findIndex(h=>norm(h)==='SABER_4'||norm(h)==='SABER 4');
    const iCS=headers.findIndex(h=>norm(h)==='CONS_SABER'||norm(h)==='CONSOLIDADO_SABER'||norm(h)==='SABER');
    const iH1=headers.findIndex(h=>norm(h)==='HACER_1'||norm(h)==='HACER 1');
    const iH2=headers.findIndex(h=>norm(h)==='HACER_2'||norm(h)==='HACER 2');
    const iH3=headers.findIndex(h=>norm(h)==='HACER_3'||norm(h)==='HACER 3');
    const iH4=headers.findIndex(h=>norm(h)==='HACER_4'||norm(h)==='HACER 4');
    const iCH=headers.findIndex(h=>norm(h)==='CONS_HACER'||norm(h)==='CONSOLIDADO_HACER'||norm(h)==='HACER');
    const iSer1=headers.findIndex(h=>norm(h)==='SER_1'||norm(h)==='SER 1');
    const iSer2=headers.findIndex(h=>norm(h)==='SER_2'||norm(h)==='SER 2');
    const iCSer=headers.findIndex(h=>norm(h)==='CONS_SER'||norm(h)==='CONSOLIDADO_SER'||norm(h)==='SER');

    const toN = (v) => parseFloat(String(v||'').trim().replace(',','.'));
    const asig      = iAsig >= 0 ? String(row[iAsig]||'').trim() : 'Asignatura';
    const prom      = iProm >= 0 ? toN(row[iProm]) : NaN;
    const rec       = iRec  >= 0 ? toN(row[iRec])  : NaN;
    const consSaber = iCS   >= 0 ? toN(row[iCS])   : NaN;
    const consHacer = iCH   >= 0 ? toN(row[iCH])   : NaN;
    const consSer   = iCSer >= 0 ? toN(row[iCSer]) : NaN;
    const nivel  = nivelNota(prom);
    const emoji  = emojiAsig(asig);
    const desemp = desempenoLabel(prom);
    const nv = (idx) => idx >= 0 && String(row[idx]||'').trim() !== '' ? (isNaN(toN(row[idx])) ? '—' : toN(row[idx]).toFixed(1)) : '—';
    const nb = (val) => isNaN(val) ? '—' : val.toFixed(1);
    const nc = (val) => isNaN(val) ? 'sin' : nivelNota(val);

    html += `
      <div class="nota-card-wrap" id="wrap-${i}">
        <div class="nota-card" onclick="toggleDetail(${i})">
          <div class="nota-icon ${nivel}">${emoji}</div>
          <div class="nota-info">
            <div class="nota-asignatura">${asig}</div>
            <div class="nota-desempeno">${desemp}${!isNaN(rec)&&rec>0?' · 🔄 Rec: '+rec.toFixed(1):''}</div>
          </div>
          <div class="nota-badge ${nivel}">${isNaN(prom)?'—':prom.toFixed(1)}</div>
          <button class="nota-detail-btn" aria-label="Ver detalle">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
              <polyline id="arrow-${i}" points="6 9 12 15 18 9"/>
            </svg>
          </button>
        </div>
        <div class="nota-detail" id="detail-${i}">
          <div class="pilar-label">📘 Saber (35%)</div>
          <div class="detail-grid detail-grid-4">
            <div class="detail-item"><div class="detail-item-label">S1</div><div class="detail-item-val">${nv(iS1)}</div></div>
            <div class="detail-item"><div class="detail-item-label">S2</div><div class="detail-item-val">${nv(iS2)}</div></div>
            <div class="detail-item"><div class="detail-item-label">S3</div><div class="detail-item-val">${nv(iS3)}</div></div>
            <div class="detail-item"><div class="detail-item-label">S4</div><div class="detail-item-val">${nv(iS4)}</div></div>
          </div>
          <div class="pilar-cons ${nc(consSaber)}">Consolidado Saber: <strong>${nb(consSaber)}</strong></div>
          <div class="pilar-label" style="margin-top:10px">📗 Hacer (35%)</div>
          <div class="detail-grid detail-grid-4">
            <div class="detail-item"><div class="detail-item-label">H1</div><div class="detail-item-val">${nv(iH1)}</div></div>
            <div class="detail-item"><div class="detail-item-label">H2</div><div class="detail-item-val">${nv(iH2)}</div></div>
            <div class="detail-item"><div class="detail-item-label">H3</div><div class="detail-item-val">${nv(iH3)}</div></div>
            <div class="detail-item"><div class="detail-item-label">H4</div><div class="detail-item-val">${nv(iH4)}</div></div>
          </div>
          <div class="pilar-cons ${nc(consHacer)}">Consolidado Hacer: <strong>${nb(consHacer)}</strong></div>
          <div class="pilar-label" style="margin-top:10px">📙 Ser (30%)</div>
          <div class="detail-grid detail-grid-4">
            <div class="detail-item"><div class="detail-item-label">Ser 1</div><div class="detail-item-val">${nv(iSer1)}</div></div>
            <div class="detail-item"><div class="detail-item-label">Ser 2</div><div class="detail-item-val">${nv(iSer2)}</div></div>
          </div>
          <div class="pilar-cons ${nc(consSer)}">Consolidado Ser: <strong>${nb(consSer)}</strong></div>
          <div class="prom-final-row">
            <span>Promedio del periodo</span>
            <span class="nota-badge ${nivel}" style="font-size:18px;padding:4px 14px">${isNaN(prom)?'—':prom.toFixed(1)}</span>
          </div>
        </div>
      </div>`;
  });

  container.innerHTML = html;
}

window.toggleDetail = function(i) {
  const detail = document.getElementById('detail-' + i);
  const wrap   = document.getElementById('wrap-' + i);
  const arrow  = document.getElementById('arrow-' + i);
  const open   = detail.classList.toggle('open');
  wrap.classList.toggle('expanded', open);
  if (arrow) arrow.setAttribute('points', open ? '6 15 12 9 18 15' : '6 9 12 15 18 9');
};

// ════════════════════════════════════════════════════════
//  RENDER — ASISTENCIA
// ════════════════════════════════════════════════════════
const NOMBRES_ASIG = {
  CCNN:'Ciencias Naturales', FISICA:'Física', QUIMICA:'Química',
  CCSS:'Ciencias Sociales', CATEDRA:'Cátedra para la Paz',
  ARTISTICA:'Ed. Artística', ETICA:'Ética y Valores', EDUFISICA:'Ed. Física',
  RELIGION:'Ed. Religiosa', LENGUAJE:'Lengua Castellana', INGLES:'Inglés',
  MATEMATICAS:'Matemáticas', ESTADISTICA:'Estadística', GEOMETRIA:'Geometría',
  TECNOLOGIA:'Tecnología e Informática', FILOSOFIA:'Filosofía',
  ECONOMICAS:'Ciencias Económicas', EMPRENDIMIENTO:'Emprendimiento',
  CI:'Centro de Interés', ALGEBRA:'Álgebra', COMPORTAMIENTO:'Comportamiento',
};
const LIMITE_FALLAS = 15;

function renderAsistencia() {
  const container = document.getElementById('asistencia-list');
  const periodo   = G.periodoSelect;

  if (G.asistencia.length === 0) {
    container.innerHTML = emptyHTML('📅', 'No se encontraron registros de asistencia.');
    activarSemaforo('verde', '¡Sin faltas registradas!', 'Todo en regla por ahora');
    return;
  }

  const { headers, row } = G.asistencia[0];
  const asignaturas = [];
  Object.keys(NOMBRES_ASIG).forEach(codigo => {
    const idx = headers.findIndex(h => norm(h) === norm(codigo + '_P' + periodo));
    if (idx < 0) return;
    const fallas = parseFloat(String(row[idx]||'').trim().replace(',','.'));
    if (isNaN(fallas)) return;
    asignaturas.push({ codigo, nombre: NOMBRES_ASIG[codigo], fallas });
  });

  if (asignaturas.length === 0) {
    container.innerHTML = emptyHTML('📅', `No hay registros de asistencia para el Periodo ${periodo}.`);
    activarSemaforo('verde', 'Sin datos aún', 'El docente aún no ha registrado asistencia');
    return;
  }

  const cuentaNormal  = asignaturas.filter(a => (a.fallas/LIMITE_FALLAS) < 0.5).length;
  const cuentaRiesgo  = asignaturas.filter(a => (a.fallas/LIMITE_FALLAS) >= 0.5 && (a.fallas/LIMITE_FALLAS) < 0.8).length;
  const cuentaCritico = asignaturas.filter(a => (a.fallas/LIMITE_FALLAS) >= 0.8).length;
  const sub = (n) => n > 0 ? `(${n} asignatura${n!==1?'s':''})` : '';
  document.getElementById('sem-normal-sub').textContent  = sub(cuentaNormal);
  document.getElementById('sem-riesgo-sub').textContent  = sub(cuentaRiesgo);
  document.getElementById('sem-critico-sub').textContent = sub(cuentaCritico);

  if (cuentaCritico > 0) {
    activarSemaforo('rojo','¡Riesgo de pérdida!',`${cuentaCritico} asignatura${cuentaCritico!==1?'s':''} en estado crítico`);
  } else if (cuentaRiesgo > 0) {
    activarSemaforo('amarillo','¡Atención!',`${cuentaRiesgo} asignatura${cuentaRiesgo!==1?'s':''} en riesgo`);
  } else {
    activarSemaforo('verde','¡Buena asistencia!','Todas tus asignaturas están en regla');
  }

  const conFallas = asignaturas.filter(a => a.fallas > 0).sort((a,b) => b.fallas-a.fallas);
  const sinFallas = asignaturas.filter(a => a.fallas === 0);
  let html = '';
  if (conFallas.length > 0) {
    html += `<div class="asist-seccion-label">⚠️ Con fallas registradas</div>`;
    conFallas.forEach(({ nombre, fallas }) => { html += tarjetaFalla(nombre, fallas); });
  }
  if (sinFallas.length > 0) {
    html += `<div class="asist-seccion-label" style="margin-top:${conFallas.length?'16px':'0'}">✅ Sin fallas</div>`;
    sinFallas.forEach(({ nombre }) => { html += tarjetaFalla(nombre, 0); });
  }
  container.innerHTML = html;
}

function tarjetaFalla(nombre, fallas) {
  const pct = Math.min((fallas/LIMITE_FALLAS)*100, 100);
  let colorBar='verde', chipClass='ok';
  if (fallas===0)    { colorBar='verde';    chipClass='ok';     }
  else if (pct < 50) { colorBar='verde';    chipClass='ok';     }
  else if (pct < 80) { colorBar='amarillo'; chipClass='alerta'; }
  else               { colorBar='rojo';     chipClass='fallas'; }
  return `
    <div class="asist-card">
      <div class="asist-header">
        <span class="asist-nombre">${nombre}</span>
        <span class="asist-chip ${fallas===0?'ok':chipClass}">${fallas} falla${fallas!==1?'s':''}</span>
      </div>
      <div class="asist-bar-bg">
        <div class="asist-bar-fill ${colorBar}" style="width:${pct}%"></div>
      </div>
    </div>`;
}

function activarSemaforo(color, status, detail) {
  document.getElementById('sem-normal').classList.remove('activo');
  document.getElementById('sem-riesgo').classList.remove('activo');
  document.getElementById('sem-critico').classList.remove('activo');
  const mapa = { verde:'sem-normal', amarillo:'sem-riesgo', rojo:'sem-critico' };
  if (mapa[color]) document.getElementById(mapa[color]).classList.add('activo');
  const msg = document.getElementById('semaforo-mensaje');
  msg.textContent = status + (detail?' — '+detail:'');
  msg.className   = 'semaforo-mensaje ' + color;
}

// ════════════════════════════════════════════════════════
//  RENDER — RECONOCIMIENTOS
// ════════════════════════════════════════════════════════
function renderReconocimientos() {
  const container = document.getElementById('reconocimientos-list');
  let html = '';
  const felicitaciones = G.observaciones.filter(o => o.tipo === 'Felicitación');
  felicitaciones.slice().reverse().forEach(({ headers, row }, i) => {
    const iPeriodo = headers.findIndex(h => norm(h).includes('PERIODO'));
    const iObs     = headers.findIndex(h => norm(h).includes('OBSERVA'));
    const iFecha   = headers.findIndex(h => norm(h).includes('FECHA'));
    const iDocente = headers.findIndex(h => norm(h).includes('DOCENTE'));
    const iAsig    = headers.findIndex(h => norm(h).includes('ASIGNATURA'));
    const periodo  = String(row[iPeriodo]||'').trim();
    const mensaje  = String(row[iObs]    ||'').trim();
    const fecha    = String(row[iFecha]  ||'').trim();
    const docente  = String(row[iDocente]||'').trim();
    const asig     = String(row[iAsig]   ||'').trim();
    html += `
      <div class="honor-card" style="animation-delay:${i*0.05}s">
        <div class="honor-header">
          <div class="honor-emoji">🏆</div>
          <div>
            <div class="honor-title">Felicitación${asig?' — '+asig:''}</div>
            ${periodo?`<div class="honor-period">Periodo ${periodo}</div>`:''}
          </div>
        </div>
        <div class="honor-msg">${mensaje}</div>
        <div class="honor-footer">${docente?'👤 '+docente:''}${fecha?' · 📅 '+fecha:''}</div>
      </div>`;
  });
  G.avisos.forEach(({ headers, row }) => {
    const iTitle = headers.findIndex(h => norm(h).includes('TITULO'));
    const iMsg   = headers.findIndex(h => norm(h).includes('MENSAJE')||norm(h).includes('CONTENIDO'));
    const iFecha = headers.findIndex(h => norm(h).includes('FECHA'));
    const title  = String(row[iTitle]||'').trim();
    const msg    = String(row[iMsg]  ||'').trim();
    const fecha  = String(row[iFecha]||'').trim();
    html += `
      <div class="aviso-card">
        <div class="aviso-title">📢 ${title}</div>
        <div class="aviso-msg">${msg}</div>
        ${fecha?`<div class="aviso-meta">${fecha}</div>`:''}
      </div>`;
  });
  if (html==='') html = emptyHTML('🏅','¡Aún no tienes felicitaciones registradas. Sigue esforzándote!');
  container.innerHTML = html;
}

// ════════════════════════════════════════════════════════
//  RENDER — OBSERVACIONES
// ════════════════════════════════════════════════════════
function renderObservaciones() {
  const container = document.getElementById('observaciones-list');
  const soloObs   = G.observaciones.filter(o => o.tipo !== 'Felicitación');
  if (soloObs.length === 0) {
    container.innerHTML = emptyHTML('📝','No tienes anotaciones registradas por docentes. ¡Eso es buena señal!');
    return;
  }
  let html = '';
  [...soloObs].reverse().forEach(({ headers, row }, i) => {
    const iPeriodo = headers.findIndex(h => norm(h).includes('PERIODO'));
    const iObs     = headers.findIndex(h => norm(h).includes('OBSERVACION')||norm(h).includes('ANOTACION'));
    const iFecha   = headers.findIndex(h => norm(h).includes('FECHA'));
    const iDocente = headers.findIndex(h => norm(h).includes('DOCENTE')||norm(h).includes('PROFESOR'));
    const iAsig    = headers.findIndex(h => norm(h).includes('ASIGNATURA')||norm(h).includes('MATERIA'));
    const periodo  = String(row[iPeriodo]||'').trim();
    const obs      = String(row[iObs]    ||'').trim();
    const fecha    = String(row[iFecha]  ||'').trim();
    const docente  = String(row[iDocente]||'').trim();
    const asig     = String(row[iAsig]   ||'').trim();
    if (!obs) return;
    html += `
      <div class="obs-card" style="animation-delay:${i*0.05}s">
        <div class="obs-header">
          ${periodo?`<span class="obs-periodo">Periodo ${periodo}</span>`:''}
          ${asig?`<span class="obs-asignatura">${asig}</span>`:''}
        </div>
        <div class="obs-texto">${obs}</div>
        <div class="obs-footer">
          <span>${docente?'👤 '+docente:''}</span>
          <span>${fecha?'📅 '+fecha:''}</span>
        </div>
      </div>`;
  });
  container.innerHTML = html || emptyHTML('📝','No hay anotaciones para mostrar.');
}

// ════════════════════════════════════════════════════════
//  PERFIL
// ════════════════════════════════════════════════════════
function renderPerfil() {
  const card = document.getElementById('perfil-info-card');
  if (!card) return;
  card.innerHTML = `
    <div class="perfil-card">
      <div class="perfil-avatar-row">
        <div class="perfil-avatar">${(G.nombre[0]||'?').toUpperCase()}</div>
        <div>
          <div class="perfil-nombre">${G.nombre}</div>
          <span class="perfil-rol">🎒 Estudiante</span>
        </div>
      </div>
      <div class="perfil-datos">
        <div class="perfil-dato-row">
          <span class="perfil-dato-label">Grado</span>
          <span class="perfil-dato-valor">${G.grado||'—'}</span>
        </div>
        <div class="perfil-dato-row">
          <span class="perfil-dato-label">Sede</span>
          <span class="perfil-dato-valor">${G.sede||'—'}</span>
        </div>
        <div class="perfil-dato-row">
          <span class="perfil-dato-label">Usuario (documento)</span>
          <span class="perfil-dato-valor">${G.id||'—'}</span>
        </div>
      </div>
      <div class="perfil-aviso">
        🔒 Tu información personal solo puede ser modificada por la institución.
        Si hay un error en tus datos, comunícate con secretaría.
      </div>
    </div>`;

  const pwdCard = document.querySelector('.cambiar-pwd-card');
  if (pwdCard && !pwdCard.querySelector('.cambiar-pwd-header')) {
    const oldTitle = pwdCard.querySelector('.cambiar-pwd-title');
    if (oldTitle) {
      oldTitle.outerHTML = `
        <div class="cambiar-pwd-header">
          <div class="cambiar-pwd-icon">🔐</div>
          <div>
            <div class="cambiar-pwd-title">Cambiar contraseña</div>
            <div class="cambiar-pwd-sub">Tu contraseña inicial es tu fecha de nacimiento sin barras. Ej: <code style="font-family:monospace;background:#f0f2f7;padding:1px 6px;border-radius:5px;">17022010</code></div>
          </div>
        </div>`;
    }
  }
}

// ════════════════════════════════════════════════════════
//  CAMBIAR CONTRASEÑA
// ════════════════════════════════════════════════════════
window.togglePwd = function(id) {
  const inp = document.getElementById(id);
  inp.type  = inp.type === 'password' ? 'text' : 'password';
};

window.cambiarContrasena = async function() {
  const actual    = document.getElementById('pwd-actual').value.trim();
  const nueva     = document.getElementById('pwd-nueva').value.trim();
  const confirmar = document.getElementById('pwd-confirmar').value.trim();
  const errEl     = document.getElementById('pwd-error');
  const okEl      = document.getElementById('pwd-success');
  const btnEl     = document.querySelector('.btn-cambiar-pwd');

  errEl.classList.remove('show');
  okEl.classList.remove('show');

  if (!actual||!nueva||!confirmar) { errEl.textContent='⚠️ Completa todos los campos.'; errEl.classList.add('show'); return; }
  if (nueva.length < 6)            { errEl.textContent='⚠️ La nueva contraseña debe tener al menos 6 caracteres.'; errEl.classList.add('show'); return; }
  if (nueva !== confirmar)         { errEl.textContent='⚠️ Las contraseñas nuevas no coinciden.'; errEl.classList.add('show'); return; }

  try {
    btnEl.textContent = 'Verificando…';
    btnEl.disabled    = true;

    const url  = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent('asignaturas_estudiantes')}?key=${API_KEY}`;
    const res  = await fetch(url);
    const data = await res.json();
    if (!data.values || data.values.length < 2) throw new Error('Sin datos');

    const headers = data.values[0];
    const iId    = headers.findIndex(h => normH(h) === 'ID_ESTUDIANTE' || normH(h).includes('IDENTIFICADOR'));
    const iFecha = headers.findIndex(h => normH(h).includes('FECHA') && normH(h).includes('NACIMIENTO'));
    const iPwd   = headers.findIndex(h => normH(h) === 'CONTRASENA' || normH(h) === 'CONTRASEÑA');

    let claveActual = '', encontrado = false;
    for (let r = 1; r < data.values.length; r++) {
      const row = data.values[r];
      if (String(row[iId]||'').trim() === G.id) {
        const claveBase     = normalizarFechaLocal(String(row[iFecha]||'').trim());
        const clavePersonal = iPwd >= 0 ? String(row[iPwd]||'').trim() : '';
        claveActual = clavePersonal || claveBase;
        encontrado  = true;
        break;
      }
    }

    if (!encontrado) { errEl.textContent='⚠️ No se encontró tu cuenta.'; errEl.classList.add('show'); btnEl.textContent='Guardar nueva contraseña'; btnEl.disabled=false; return; }
    if (actual !== claveActual) { errEl.textContent='❌ La contraseña actual es incorrecta.'; errEl.classList.add('show'); btnEl.textContent='Guardar nueva contraseña'; btnEl.disabled=false; return; }

    btnEl.textContent = 'Guardando…';
    const postRes  = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      body:   JSON.stringify({ tipo:'cambiarContrasenaEstudiante', id:G.id, nombre:G.nombre, grado:G.grado, nuevaClave:nueva })
    });
    const postData = await postRes.json();

    if (postData.status === 'ok') {
      okEl.classList.add('show');
      document.getElementById('pwd-actual').value    = '';
      document.getElementById('pwd-nueva').value     = '';
      document.getElementById('pwd-confirmar').value = '';
      toast('✅ Contraseña actualizada');
    } else {
      errEl.textContent = '⚠️ ' + (postData.mensaje||'No se pudo guardar.');
      errEl.classList.add('show');
    }
  } catch(e) {
    errEl.textContent = '⚠️ Error de conexión. Verifica tu internet.';
    errEl.classList.add('show');
  }

  btnEl.textContent = 'Guardar nueva contraseña';
  btnEl.disabled    = false;
};

// ── LOGOUT ──
document.getElementById('btn-logout').addEventListener('click', () => {
  if (confirm('¿Seguro que quieres cerrar sesión?')) {
    sessionStorage.clear();
    localStorage.clear();
    window.location.href = 'Index.html';
  }
});