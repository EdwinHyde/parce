// ════════════════════════════════════════════════════════
//  CONFIGURACIÓN
// ════════════════════════════════════════════════════════
const API_KEY        = "AIzaSyA1GsdvskVoIQDSOg6EI8aEQW5MnQsFeoE";
const SPREADSHEET_ID = "1ytP_bXfbjRsA05KJYM8HAxQvTvnAeHgz09r4QhMy-Rs";
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbygBydUcKQWczkcs4jL3BshWNY1npLtsqK9IbNA72nIgpRhmjHo3tKnCVry0AcwvGIf/exec";

const SHEETS = {
  usuarios:    "Usuarios",
  estudiantes: "asignaturas_estudiantes",
  periodos:    "Periodos",
  consolidado: "Consolidado",
  avisos:      "Avisos",
  configBoletines: 'Configuracion_boletines'
};

// ════════════════════════════════════════════════════════
//  ESTADO GLOBAL
// ════════════════════════════════════════════════════════
let nombreAdmin = "";
let sedeAdmin   = "";
let dataCache   = {};

// ════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════
function getUrlParameter(name) {
  name = name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]');
  const regex = new RegExp('[\\?&]' + name + '=([^&#]*)');
  const results = regex.exec(location.search);
  return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
}

function fetchSheet(sheetName, forceRefresh) {
  if (!forceRefresh && dataCache[sheetName]) return Promise.resolve(dataCache[sheetName]);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(sheetName)}?key=${API_KEY}`;
  return fetch(url).then(r => r.json()).then(data => {
    dataCache[sheetName] = data;
    return data;
  });
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

function colorNota(val) {
  const n = parseFloat(val);
  if (isNaN(n)) return '';
  if (n >= 3.5) return 'nota-alta';
  if (n >= 3.0) return 'nota-media';
  return 'nota-baja';
}


function norm(h) {
  return String(h||'').trim().toUpperCase()
    .replace(/[ÁÀÂÄ]/g,'A').replace(/[ÉÈÊË]/g,'E')
    .replace(/[ÍÌÎÏ]/g,'I').replace(/[ÓÒÔÖ]/g,'O').replace(/[ÚÙÛÜ]/g,'U')
    .replace(/[áàâä]/g,'A').replace(/[éèêë]/g,'E')
    .replace(/[íìîï]/g,'I').replace(/[óòôö]/g,'O').replace(/[úùûü]/g,'U');
}

function showMsg(id, msg, tipo) {
  const el = document.getElementById(id);
  if (!el) return;
  
  el.innerHTML = msg;                    // ← Cambiado a innerHTML
  el.className = `msg msg-${tipo} visible`;
  
  setTimeout(() => {
    el.classList.remove('visible');
  }, 4000);
}

function enviarAppsScript(payload) {
  return new Promise((resolve, reject) => {
    fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(payload),
      redirect: 'follow'
    })
    .then(r => r.text().then(t => {
      try { resolve(JSON.parse(t)); }
      catch { resolve({ status:'ok', mensaje:'Operación completada.' }); }
    }))
    .catch(reject);
  });
}

// Orden estándar de grados
const ORDEN_GRADOS = ['Preescolar','Primero','Segundo','Tercero','Cuarto','Quinto',
                      'Sexto','Séptimo','Octavo','Noveno','Décimo','Undécimo'];

function ordenarGrados(grados) {
  return [...grados].sort((a, b) => {
    const ia = ORDEN_GRADOS.indexOf(a);
    const ib = ORDEN_GRADOS.indexOf(b);
    if (ia === -1 && ib === -1) return a.localeCompare(b);
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });
}

// ════════════════════════════════════════════════════════
//  LOGOUT
// ════════════════════════════════════════════════════════
document.getElementById('btn-logout').addEventListener('click', () => {
  if (confirm('¿Cerrar sesión?')) window.location.href = 'index.html';
});

// ════════════════════════════════════════════════════════
//  MENÚ MÓVIL
// ════════════════════════════════════════════════════════
document.getElementById('btn-menu-toggle').addEventListener('click', () => {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebar-overlay').classList.toggle('open');
});

document.getElementById('sidebar-overlay').addEventListener('click', () => {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebar-overlay').classList.remove('open');
});

// ════════════════════════════════════════════════════════
//  MÓDULO: DASHBOARD
// ════════════════════════════════════════════════════════

// Cache de datos del dashboard para no re-fetchar al cambiar periodo
let _dashboardData = null;
let _dashboardPeriodos = [];

async function cargarDashboard() {
  // Limpiar cache al entrar al dashboard
  _dashboardData = null;
  _dashboardPeriodos = [];

  // Mostrar spinners
  ['dash-semaforo','dash-barras','dash-top3'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = '<div class="module-loading"><div class="ring-loader"></div><p>Cargando...</p></div>';
  });

  try {
    const [dataEst, dataPer, dataUs, dataCalif, dataAsist] = await Promise.all([
      fetchSheet(SHEETS.estudiantes),
      fetchSheet(SHEETS.periodos),
      fetchSheet(SHEETS.usuarios),
      fetchSheet('Calificaciones'),
      fetchSheet('Asistencia')
    ]);

    // Guardar en cache
    _dashboardData = { dataEst, dataPer, dataUs, dataCalif, dataAsist };

    // ── 1. Procesar periodos disponibles ─────────────────
    const hoy = new Date(); hoy.setHours(0,0,0,0);
    _dashboardPeriodos = [];

    if (!dataPer.error && dataPer.values?.length > 1) {
      const hPer  = dataPer.values[0].map(h => h.trim());
      const iNum  = hPer.findIndex(h => norm(h).includes('PERIODO'));
      const iIni  = hPer.findIndex(h => norm(h).includes('INICIO'));
      const iFin  = hPer.findIndex(h => (norm(h).includes('FINAL') || norm(h).includes('FIN')) && !norm(h).includes('CIERRE') && !norm(h).includes('SISTEMA'));
      const iCie  = hPer.findIndex(h => norm(h).includes('CIERRE') && norm(h).includes('SISTEMA'));
      const iFinA = hPer.findIndex(h => (norm(h).includes('FINAL') || norm(h).includes('FIN')) && !norm(h).includes('CIERRE'));

      dataPer.values.slice(1).forEach(row => {
        const num = String(row[iNum]||'').trim();
        if (!num) return;
        const ini  = iIni  >= 0 ? parseFecha(row[iIni])  : null;
        const fin  = iFin  >= 0 ? parseFecha(row[iFin])  : (iFinA >= 0 ? parseFecha(row[iFinA]) : null);
        const cie  = iCie  >= 0 ? parseFecha(row[iCie])  : null;
        const activo = ini && fin && hoy >= ini && hoy <= fin;
        _dashboardPeriodos.push({ num, ini, fin, cie, activo });
      });
    }

    // ── 2. Construir selector de periodos ─────────────────
    const sel = document.getElementById('dash-select-periodo');
    if (sel && _dashboardPeriodos.length) {
      sel.innerHTML = '';
      // Detectar periodo activo o usar el último
      let periodoDefault = _dashboardPeriodos.find(p => p.activo);
      if (!periodoDefault) periodoDefault = _dashboardPeriodos[_dashboardPeriodos.length - 1];

      _dashboardPeriodos.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.num;
        opt.textContent = p.activo ? `Periodo ${p.num} (actual)` : `Periodo ${p.num}`;
        if (p.num === periodoDefault?.num) opt.selected = true;
        sel.appendChild(opt);
      });

      document.getElementById('dash-periodo-controles').style.display = 'flex';
      document.getElementById('dash-periodo-controles').style.alignItems = 'center';
      document.getElementById('dash-periodo-controles').style.flexWrap = 'wrap';
      document.getElementById('dash-periodo-controles').style.gap = '8px';

      sel.addEventListener('change', () => {
        renderizarDashboard(sel.value);
      });

      document.getElementById('dash-btn-excel').addEventListener('click', () => {
        exportarExcelPeriodo(sel.value);
      });

      await renderizarDashboard(periodoDefault?.num || _dashboardPeriodos[0]?.num);
    } else {
      await renderizarDashboard(null);
    }

  } catch(err) {
    console.error('[Dashboard Admin] Error:', err);
    ['dash-semaforo','dash-barras','dash-top3'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.innerHTML = '<div class="empty-state"><p>Error al cargar datos. Revisa la consola.</p></div>';
    });
  }
}

async function renderizarDashboard(numPeriodoSel) {
  if (!_dashboardData) return;
  const { dataEst, dataPer, dataCalif, dataAsist } = _dashboardData;
  const hoy = new Date(); hoy.setHours(0,0,0,0);

  const periodoInfo = _dashboardPeriodos.find(p => p.num === String(numPeriodoSel)) || null;
  const numPeriodo  = numPeriodoSel ? String(numPeriodoSel).trim() : null;

  // ── Banner ────────────────────────────────────────────
  const cardTxt = document.getElementById('dash-periodo-titulo');
  const cardSub = document.getElementById('dash-periodo-sub');
  const badges  = document.getElementById('dash-periodo-badges');

  if (periodoInfo) {
    const esActual  = periodoInfo.activo;
    const esPasado  = periodoInfo.fin && hoy > periodoInfo.fin;
    cardTxt.textContent = esActual
      ? `Periodo ${periodoInfo.num} en curso`
      : esPasado
        ? `Periodo ${periodoInfo.num} — Finalizado`
        : `Periodo ${periodoInfo.num}`;
    cardSub.textContent = periodoInfo.ini && periodoInfo.fin
      ? `${formatFecha(periodoInfo.ini)} — ${formatFecha(periodoInfo.fin)}`
      : 'Fechas no configuradas';
    const abierto = periodoInfo.cie ? hoy <= periodoInfo.cie : esActual;
    badges.innerHTML = esActual
      ? `<span class="periodo-badge ${abierto?'open':'closed'}">${abierto?'Sistema Habilitado':'Sistema cerrado'}</span>
         ${periodoInfo.cie ? `<span class="periodo-badge">Límite: ${formatFecha(periodoInfo.cie)}</span>` : ''}`
      : esPasado
        ? '<span class="periodo-badge closed">Periodo cerrado</span>'
        : '<span class="periodo-badge">Próximo</span>';
  } else {
    cardTxt.textContent = 'Sin periodos configurados';
    cardSub.textContent = 'Ve al módulo de Periodos para agregar uno';
    badges.innerHTML = '';
  }

  // Actualizar etiquetas dinámicas
  const labelPer = numPeriodo ? `Periodo ${numPeriodo}` : 'Periodo actual';
  const elLabelSem = document.getElementById('dash-label-periodo-semaforo');
  const elLabelBar = document.getElementById('dash-label-periodo-barras');
  if (elLabelSem) elLabelSem.textContent = labelPer;
  if (elLabelBar) elLabelBar.textContent = labelPer;

  // ── Verificar si el botón Excel debe mostrarse ─────────
  // Solo se muestra si el periodo ya finalizó (fecha fin < hoy)
  const btnExcel = document.getElementById('dash-btn-excel');
  if (btnExcel) {
    const periodoFinalizado = periodoInfo && periodoInfo.fin && hoy > periodoInfo.fin;
    btnExcel.style.display = periodoFinalizado ? 'flex' : 'none';
  }

  // ── Mapa de nombres y grados ─────────────────────────
  const mapaNombres = {};
  const mapaGrado   = {};
  const mapaGradoInfo = {}; // grado → { nombre, estado }
  if (!dataEst.error && dataEst.values?.length > 1) {
    const hE  = dataEst.values[0].map(h => h.trim());
    const iNm = hE.findIndex(h => norm(h) === 'NOMBRES_APELLIDOS' || norm(h).includes('NOMBRE'));
    const iId = hE.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
    const iGr = hE.findIndex(h => norm(h) === 'GRADO');
    dataEst.values.slice(1).forEach(r => {
      const id = (r[iId]||'').trim();
      if (id) {
        mapaNombres[id] = (r[iNm]||'').trim();
        mapaGrado[id]   = (r[iGr]||'').trim();
      }
    });
  }

  // ── Análisis de Calificaciones ────────────────────────
  const enRiesgoIds      = new Set();
  const bajosPorEst      = {};
  const promediosPorEst  = {};
  const promediosPorGrado= {};

  if (!dataCalif.error && dataCalif.values?.length > 1 && numPeriodo) {
    const hC    = dataCalif.values[0].map(h => h.trim());
    const iP    = hC.findIndex(h => norm(h) === 'PERIODO');
    const iId   = hC.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
    const iProm = hC.findIndex(h => norm(h).includes('PROMEDIO'));
    const iDesp = hC.findIndex(h => norm(h).includes('DESEMPE'));
    const iGr   = hC.findIndex(h => norm(h) === 'GRADO');

    dataCalif.values.slice(1).forEach(row => {
      const rowPer = String(row[iP]||'').trim().replace(/\.0$/,'');
      const numPer = numPeriodo.replace(/\.0$/,'');
      if (rowPer !== numPer) return;
      const id   = (row[iId]||'').trim();
      if (!id) return;
      const prom = parseFloat(String(row[iProm]||'').replace(',','.'));
      const desp = norm(row[iDesp]||'');
      const grad = (row[iGr]||'').trim() || mapaGrado[id] || '';

      if (!isNaN(prom) && prom > 0) {
        if (!promediosPorEst[id]) promediosPorEst[id] = { suma:0, count:0, grado:grad };
        promediosPorEst[id].suma  += prom;
        promediosPorEst[id].count += 1;
        if (prom < 3.0) enRiesgoIds.add(id);
        if (grad) {
          if (!promediosPorGrado[grad]) promediosPorGrado[grad] = [];
          promediosPorGrado[grad].push(prom);
        }
      }
      if (desp === 'BAJO') bajosPorEst[id] = (bajosPorEst[id]||0) + 1;
    });
  }

  // ── Tarjeta 1: Estudiantes en riesgo ─────────────────
  document.getElementById('dash-en-riesgo').textContent = enRiesgoIds.size || '0';
  document.getElementById('dash-en-riesgo-sub').textContent = numPeriodo
    ? `con promedio < 3.0 en periodo ${numPeriodo}`
    : 'configura un periodo activo';

  // ── Tarjeta 3: Grado con más riesgo ──────────────────
  const riesgoPorGrado = {};
  enRiesgoIds.forEach(id => {
    const g = mapaGrado[id] || promediosPorEst[id]?.grado || '';
    if (g) riesgoPorGrado[g] = (riesgoPorGrado[g]||0) + 1;
  });
  const gradoMasRiesgo = Object.entries(riesgoPorGrado).sort((a,b) => b[1]-a[1])[0];
  if (gradoMasRiesgo) {
    document.getElementById('dash-grado-riesgo').textContent     = gradoMasRiesgo[0];
    document.getElementById('dash-grado-riesgo-sub').textContent =
      `${gradoMasRiesgo[1]} estudiante${gradoMasRiesgo[1]!==1?'s':''} en riesgo`;
  } else {
    document.getElementById('dash-grado-riesgo').textContent     = '—';
    document.getElementById('dash-grado-riesgo-sub').textContent = 'sin datos del periodo';
  }

  // ── Tarjeta 4: Más asignaturas en BAJO ───────────────
  const masRiesgo = Object.entries(bajosPorEst).sort((a,b) => b[1]-a[1])[0];
  if (masRiesgo) {
    const [idR, cantR] = masRiesgo;
    const nombreCorto  = (mapaNombres[idR]||'Sin nombre').split(' ').slice(0,2).join(' ');
    document.getElementById('dash-mas-bajo-nombre').textContent = nombreCorto;
    document.getElementById('dash-mas-bajo-sub').textContent =
      `${cantR} asignatura${cantR!==1?'s':''} en BAJO · ${mapaGrado[idR]||''}`;
  } else {
    document.getElementById('dash-mas-bajo-nombre').textContent = '—';
    document.getElementById('dash-mas-bajo-sub').textContent    = 'sin datos del periodo';
  }

  // ── Análisis de Asistencia ────────────────────────────
  const fallasPorEst = {};
  if (!dataAsist.error && dataAsist.values?.length > 1 && numPeriodo) {
    const hA   = dataAsist.values[0].map(h => (h||'').trim());
    const iP   = hA.findIndex(h => norm(h) === 'PERIODO');
    const iId  = hA.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
    const iF   = hA.findIndex(h => norm(h) === 'FALLAS');
    const numPer = numPeriodo.replace(/\.0$/,'');
    dataAsist.values.slice(1).forEach(row => {
      const rowPer = String(row[iP]||'').trim().replace(/\.0$/,'');
      if (rowPer !== numPer) return;
      const id = iId >= 0 ? (row[iId]||'').trim() : '';
      if (!id) return;
      const f = iF >= 0 ? parseInt(row[iF])||0 : 0;
      if (f > 0) fallasPorEst[id] = (fallasPorEst[id]||0) + f;
    });
  }

  // ── Tarjeta 2: Mayor ausentismo ───────────────────────
  const masFallas = Object.entries(fallasPorEst).sort((a,b) => b[1]-a[1])[0];
  if (masFallas) {
    const [idF, cantF] = masFallas;
    const nombreCorto  = (mapaNombres[idF]||'Sin nombre').split(' ').slice(0,2).join(' ');
    document.getElementById('dash-mas-fallas-nombre').textContent = nombreCorto;
    document.getElementById('dash-mas-fallas-sub').textContent =
      `${cantF} falla${cantF!==1?'s':''} · ${mapaGrado[idF]||''}`;
  } else {
    document.getElementById('dash-mas-fallas-nombre').textContent = '—';
    document.getElementById('dash-mas-fallas-sub').textContent    = 'sin registros de asistencia';
  }

  // ── Top 3 mejores promedios ───────────────────────────
  const top3 = Object.entries(promediosPorEst)
    .map(([id, d]) => ({
      id,
      nombre: mapaNombres[id] || id,
      grado:  d.grado || mapaGrado[id] || '—',
      prom:   d.count > 0 ? d.suma / d.count : 0
    }))
    .filter(e => e.prom > 0)
    .sort((a,b) => b.prom - a.prom)
    .slice(0, 3);

  const medallas = ['🥇','🥈','🥉'];
  const colTop   = ['#f59e0b','#94a3b8','#b45309'];
  document.getElementById('dash-top3').innerHTML = top3.length
    ? top3.map((e, i) => `
        <div class="top3-item">
          <div class="top3-medalla" style="color:${colTop[i]}">${medallas[i]}</div>
          <div class="top3-info">
            <div class="top3-nombre">${e.nombre}</div>
            <div class="top3-grado">${e.grado}</div>
          </div>
          <div class="top3-prom" style="color:${colTop[i]}">${e.prom.toFixed(2)}</div>
        </div>`).join('')
    : `<div class="empty-state"><p>Sin datos del periodo ${numPeriodo||''}.</p></div>`;

  // ── Semáforo por grado ────────────────────────────────
  const gradosOrdenados = ordenarGrados(Object.keys(promediosPorGrado));
  if (!gradosOrdenados.length) {
    document.getElementById('dash-semaforo').innerHTML =
      `<div class="empty-state"><p>Sin datos de calificaciones para el periodo ${numPeriodo||''}.</p></div>`;
  } else {
    document.getElementById('dash-semaforo').innerHTML = `
      <div class="semaforo-grid">
        ${gradosOrdenados.map(g => {
          const proms  = promediosPorGrado[g];
          const avg    = proms.reduce((a,b)=>a+b,0) / proms.length;
          const color  = avg >= 4.0 ? 'sem-verde' : avg >= 3.0 ? 'sem-amarillo' : 'sem-rojo';
          const etiq   = avg >= 4.0 ? 'ALTO' : avg >= 3.0 ? 'BÁSICO' : 'BAJO';
          const riesgo = proms.filter(p => p < 3.0).length;
          return `
            <div class="sem-card ${color}">
              <div class="sem-grado">${g}</div>
              <div class="sem-avg">${avg.toFixed(2)}</div>
              <div class="sem-label">${etiq}</div>
              ${riesgo > 0 ? `<div class="sem-riesgo">⚠️ ${riesgo} en riesgo</div>` : ''}
            </div>`;
        }).join('')}
      </div>`;
  }

  // ── Barras por grado ──────────────────────────────────
  if (!gradosOrdenados.length) {
    document.getElementById('dash-barras').innerHTML =
      `<div class="empty-state"><p>Sin datos para el periodo ${numPeriodo||''}.</p></div>`;
  } else {
    document.getElementById('dash-barras').innerHTML = `
      <div class="barras-wrap">
        ${gradosOrdenados.map(g => {
          const proms = promediosPorGrado[g];
          const avg   = proms.reduce((a,b)=>a+b,0) / proms.length;
          const pct   = Math.round((avg / 5.0) * 100);
          const color = avg >= 4.0 ? '#16a34a' : avg >= 3.0 ? '#d97706' : '#dc2626';
          return `
            <div class="barra-row">
              <div class="barra-label">${g}</div>
              <div class="barra-track">
                <div class="barra-fill" style="width:${pct}%;background:${color};"></div>
              </div>
              <div class="barra-valor" style="color:${color};">${avg.toFixed(2)}</div>
            </div>`;
        }).join('')}
        <div class="barra-leyenda">
          <span style="color:#16a34a;">● ALTO ≥4.0</span>
          <span style="color:#d97706;">● BÁSICO ≥3.0</span>
          <span style="color:#dc2626;">● BAJO &lt;3.0</span>
        </div>
      </div>`;
  }
}

// ════════════════════════════════════════════════════════
//  EXPORTAR EXCEL — RESUMEN DE PERIODO
// ════════════════════════════════════════════════════════
async function exportarExcelPeriodo(numPeriodo) {
  const btn = document.getElementById('dash-btn-excel');
  const textoOriginal = btn ? btn.innerHTML : '';
  if (btn) { btn.innerHTML = '⏳ Generando...'; btn.disabled = true; }

  try {
    const { dataEst, dataCalif, dataAsist } = _dashboardData;
    if (!dataEst || !dataCalif) throw new Error('Datos no disponibles');

    const hE   = dataEst.values[0].map(h => (h||'').trim());
    const iNm  = hE.findIndex(h => norm(h) === 'NOMBRES_APELLIDOS' || norm(h).includes('NOMBRE'));
    const iId  = hE.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
    const iGr  = hE.findIndex(h => norm(h) === 'GRADO');
    const iSe  = hE.findIndex(h => norm(h) === 'SEDE');

    // Excluir Preescolar
    const estudiantes = dataEst.values.slice(1).filter(r => {
      const grado = norm((r[iGr]||'').trim());
      return grado !== 'PREESCOLAR' && grado !== norm('Preescolar');
    });

    // Construir mapa de calificaciones por estudiante+asignatura
    const hC    = dataCalif.values[0].map(h => (h||'').trim());
    const iCP   = hC.findIndex(h => norm(h) === 'PERIODO');
    const iCId  = hC.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
    const iCAs  = hC.findIndex(h => norm(h) === 'ASIGNATURA');
    const iCPr  = hC.findIndex(h => norm(h).includes('PROMEDIO'));
    const iCGr  = hC.findIndex(h => norm(h) === 'GRADO');

    const numPer = String(numPeriodo).replace(/\.0$/,'');
    const califMap = {}; // idEst → { asignatura → promedio }
    const asignaturasSet = new Set();

    dataCalif.values.slice(1).forEach(row => {
      const rowPer = String(row[iCP]||'').trim().replace(/\.0$/,'');
      if (rowPer !== numPer) return;
      const id   = (row[iCId]||'').trim();
      const asig = (row[iCAs]||'').trim();
      const prom = String(row[iCPr]||'').trim();
      const gr   = (row[iCGr]||'').trim();
      if (!id || !asig) return;
      // Excluir Preescolar
      if (norm(gr) === 'PREESCOLAR' || norm(gr) === norm('Preescolar')) return;
      if (!califMap[id]) califMap[id] = {};
      califMap[id][asig] = prom;
      asignaturasSet.add(asig);
    });

    // Mapa de inasistencias por estudiante
    const fallasMap = {};
    if (!dataAsist.error && dataAsist.values?.length > 1) {
      const hA  = dataAsist.values[0].map(h => (h||'').trim());
      const iAP = hA.findIndex(h => norm(h) === 'PERIODO');
      const iAI = hA.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
      const iAF = hA.findIndex(h => norm(h) === 'FALLAS');
      dataAsist.values.slice(1).forEach(row => {
        const rowPer = String(row[iAP]||'').trim().replace(/\.0$/,'');
        if (rowPer !== numPer) return;
        const id = iAI >= 0 ? (row[iAI]||'').trim() : '';
        if (!id) return;
        const f = iAF >= 0 ? parseInt(row[iAF])||0 : 0;
        fallasMap[id] = (fallasMap[id]||0) + f;
      });
    }

    const asignaturas = [...asignaturasSet].sort();

    // Construir CSV
    const encabezado = ['Documento', 'Nombre', 'Sede', 'Grado', 'Inasistencias', ...asignaturas];
    const filas = [encabezado];

    // Ordenar estudiantes por grado y nombre
    const estudiantesOrdenados = estudiantes.slice().sort((a, b) => {
      const ga = (a[iGr]||''), gb = (b[iGr]||'');
      const ia = ORDEN_GRADOS.indexOf(ga), ib = ORDEN_GRADOS.indexOf(gb);
      const cmpGrado = (ia===-1&&ib===-1) ? ga.localeCompare(gb) : (ia===-1?1:(ib===-1?-1:ia-ib));
      if (cmpGrado !== 0) return cmpGrado;
      return (a[iNm]||'').localeCompare(b[iNm]||'');
    });

    estudiantesOrdenados.forEach(r => {
      const id    = iId >= 0 ? (r[iId]||'').trim() : '';
      const nom   = iNm >= 0 ? (r[iNm]||'').trim() : '';
      const sede  = iSe >= 0 ? (r[iSe]||'').trim() : '';
      const grado = iGr >= 0 ? (r[iGr]||'').trim() : '';
      const fallas= fallasMap[id] || 0;
      const califs = asignaturas.map(a => califMap[id]?.[a] || '');
      filas.push([id, nom, sede, grado, fallas, ...califs]);
    });

    const csvContent = filas.map(f =>
      f.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',')
    ).join('\n');

    const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `Resumen_Periodo_${numPeriodo}_PARCE.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

  } catch(err) {
    console.error('Error exportando Excel:', err);
    alert('Error al generar el archivo. Revisa la consola.');
  } finally {
    if (btn) { btn.innerHTML = textoOriginal; btn.disabled = false; }
  }
}


// ════════════════════════════════════════════════════════
//  MÓDULO: USUARIOS
// ════════════════════════════════════════════════════════
let usuariosData = [];
let editandoFila = -1;

async function cargarUsuarios() {
  document.getElementById('tabla-usuarios-wrap').innerHTML =
    '<div class="module-loading"><div class="ring-loader"></div><p>Cargando usuarios...</p></div>';
  try {
    const data = await fetchSheet(SHEETS.usuarios, true);
    if (data.error || !data.values || data.values.length < 2) {
      document.getElementById('tabla-usuarios-wrap').innerHTML =
        '<div class="empty-state"><p>No hay usuarios registrados.</p></div>';
      return;
    }
    usuariosData = data.values;

    // Cargar sedes disponibles
    const dataEst = await fetchSheet(SHEETS.estudiantes);
    if (!dataEst.error && dataEst.values && dataEst.values.length > 1) {
      const hEst  = dataEst.values[0].map(h => h.trim());
      const iSede = hEst.findIndex(h => norm(h) === 'SEDE');
      window.sedesDisponibles = [...new Set(
        dataEst.values.slice(1).map(r => (r[iSede]||'').trim()).filter(Boolean)
      )].sort();
    } else {
      window.sedesDisponibles = [];
    }

    renderTablaUsuarios();
  } catch(err) {
    document.getElementById('tabla-usuarios-wrap').innerHTML =
      '<div class="empty-state"><p>Error al cargar usuarios.</p></div>';
  }
}

function renderTablaUsuarios() {
  const filtroRol = document.getElementById('filtro-rol-usuario').value;
  const headers   = usuariosData[0].map(h => h.trim());
  const rows      = usuariosData.slice(1);
  const iNombre   = headers.findIndex(h => norm(h) === 'NOMBRE');
  const iUsuario  = headers.findIndex(h => norm(h) === 'USUARIO');
  const iRol      = headers.findIndex(h => norm(h) === 'ROL');
  const iSede     = headers.findIndex(h => norm(h) === 'SEDE');
  const iEstado   = headers.findIndex(h => norm(h) === 'ESTADO' || norm(h) === 'ACTIVO');

  const filtradas = filtroRol
    ? rows.filter(r => (r[iRol]||'').trim() === filtroRol)
    : rows;

  if (!filtradas.length) {
    document.getElementById('tabla-usuarios-wrap').innerHTML =
      '<div class="empty-state"><p>Sin resultados para este filtro.</p></div>';
    return;
  }

  document.getElementById('tabla-usuarios-wrap').innerHTML = `
    <div style="overflow-x:auto;">
      <table class="data-table">
        <thead><tr>
          <th>Nombre</th><th>Usuario</th><th>Rol</th><th>Sede</th><th>Estado</th><th>Acciones</th>
        </tr></thead>
        <tbody>
          ${filtradas.map(r => {
            const rol    = (r[iRol]   ||'').trim();
            const estado = (r[iEstado]||'Activo').trim();
            const activo = estado.toLowerCase() !== 'inactivo';
            return `
            <tr>
              <td><strong>${r[iNombre]||'—'}</strong></td>
              <td style="color:var(--text-muted);font-family:monospace;font-size:12px;">${r[iUsuario]||'—'}</td>
              <td><span class="badge ${rol==='Administrador'?'badge-info':rol==='Docente'?'badge-success':'badge-gray'}">${rol||'—'}</span></td>
              <td>${r[iSede]||'—'}</td>
              <td><span class="badge ${activo?'badge-success':'badge-danger'}">${activo?'Activo':'Inactivo'}</span></td>
              <td>
                <div style="display:flex; gap:6px; flex-wrap:wrap;">
                  
                  <!-- Botón Editar - Azul suave -->
                  <button class="btn btn-sm" 
                          onclick="editarUsuario(${usuariosData.indexOf(r)})"
                          style="background: #1976d2; color: white; border: none;">
                    Editar
                  </button>
                  
                  <!-- Botón Activar / Desactivar - Naranja / Verde -->
                  <button class="btn btn-sm" 
                          onclick="toggleEstadoUsuario(${usuariosData.indexOf(r)}, '${activo ? 'Inactivo' : 'Activo'}')"
                          style="background: ${activo ? '#f57c00' : '#2e7d32'}; 
                                color: white; 
                                border: none;">
                    ${activo ? 'Desactivar' : 'Activar'}
                  </button>

                  <!-- Botón Eliminar - Rojo oscuro con hover fuerte -->
                  <button class="btn btn-sm" 
                          onclick="eliminarUsuario(${usuariosData.indexOf(r)}, '${(r[iNombre] || '').replace(/'/g, "\\'")}')"
                          title="Eliminar usuario permanentemente"
                          style="background: #c62828; color: white; border: none; padding: 5px 10px;">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                      <polyline points="3 6 5 6 21 6"></polyline>
                      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2 2 2 0 0 1 2-2"></path>
                      <line x1="10" y1="11" x2="10" y2="17"></line>
                      <line x1="14" y1="11" x2="14" y2="17"></line>
                    </svg>
                  </button>
                </div>
              </td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
}

document.getElementById('filtro-rol-usuario').addEventListener('change', renderTablaUsuarios);

document.getElementById('btn-nuevo-usuario').addEventListener('click', () => {
  editandoFila = -1;
  document.getElementById('form-usuario-title').textContent = 'Nuevo usuario';

  // Limpiar campos
  ['u-nombre', 'u-usuario', 'u-password', 'u-id', 'u-telefono'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.value = '';
      el.readOnly = false;           // Asegurar que todos sean editables
      el.style.background = 'white';
    }
  });

  // Configurar select de sede
  const selectSede = document.getElementById('u-sede');
  selectSede.innerHTML = '<option value="">Seleccionar sede</option>' +
    (window.sedesDisponibles || []).map(s => `<option value="${s}">${s}</option>`).join('');

  document.getElementById('u-rol').value    = 'Docente';
  document.getElementById('u-estado').value = 'Activo';

  document.getElementById('form-usuario-card').style.display = 'block';
  document.getElementById('form-usuario-card').scrollIntoView({ behavior: 'smooth' });
});

document.getElementById('btn-cancelar-usuario').addEventListener('click', () => {
  document.getElementById('form-usuario-card').style.display = 'none';
});


// ════════════════════════════════════════════════════════
//  EDITAR USUARIO
// ════════════════════════════════════════════════════════
function editarUsuario(filaIdx) {
  const row     = usuariosData[filaIdx];
  const headers = usuariosData[0].map(h => h.trim());
  
  const get = campo => {
    const i = headers.findIndex(h => norm(h) === norm(campo));
    return i >= 0 ? (row[i] || '').trim() : '';
  };

  editandoFila = filaIdx;
  document.getElementById('form-usuario-title').textContent = 'Editar usuario';

  document.getElementById('u-nombre').value    = get('Nombre');
  document.getElementById('u-usuario').value   = get('Usuario');
  document.getElementById('u-password').value  = get('Contraseña');
  document.getElementById('u-rol').value       = get('Rol') || 'Docente';

  // Sede
  const selectSede = document.getElementById('u-sede');
  selectSede.innerHTML = '<option value="">Seleccionar sede</option>' +
    (window.sedesDisponibles || []).map(s => `<option value="${s}">${s}</option>`).join('');
  selectSede.value = get('Sede');

  document.getElementById('u-estado').value = get('Estado') || 'Activo';

  // ID_Usuario ahora es EDITABLE
  const elId  = document.getElementById('u-id');
  const elTel = document.getElementById('u-telefono');
  if (elId) {
    elId.value = get('ID_Usuario');
    elId.readOnly = false;           // ← Ahora se puede editar
    elId.style.background = "white"; // ← Visualmente normal
  }
  if (elTel) elTel.value = get('Telefono') || get('Teléfono');

  document.getElementById('form-usuario-card').style.display = 'block';
  document.getElementById('form-usuario-card').scrollIntoView({ behavior:'smooth' });
}

function toggleEstadoUsuario(filaIdx, nuevoEstado) {
  if (!confirm(`¿Cambiar estado a "${nuevoEstado}"?`)) return;
  enviarAppsScript({ tipo: 'toggleUsuario', fila: filaIdx + 1, estado: nuevoEstado })
    .then(() => { dataCache[SHEETS.usuarios] = null; cargarUsuarios(); })
    .catch(() => alert('Error al cambiar estado.'));
}

// ════════════════════════════════════════════════════════
//  GUARDAR USUARIO
// ════════════════════════════════════════════════════════
document.getElementById('btn-guardar-usuario').addEventListener('click', async () => {
  const nombre    = document.getElementById('u-nombre').value.trim();
  const usuario   = document.getElementById('u-usuario').value.trim();
  const password  = document.getElementById('u-password').value.trim();
  const rol       = document.getElementById('u-rol').value;
  const sede      = document.getElementById('u-sede').value.trim();
  const estado    = document.getElementById('u-estado').value;
  const idUsuario = document.getElementById('u-id').value.trim();
  const telefono  = document.getElementById('u-telefono').value.trim();

  if (!nombre || !usuario) {
    showMsg('msg-usuario', 'Nombre y usuario son obligatorios.', 'error');
    return;
  }

  // Mensaje según si es creación o edición
  const esEdicion = editandoFila > 0;
  const accion = esEdicion ? 'editado' : 'creado';

  const payload = {
    tipo: 'guardarUsuario',
    fila: esEdicion ? editandoFila + 1 : -1,
    datos: {
      'ID_Usuario': idUsuario,
      'Nombre': nombre,
      'Usuario': usuario,
      'Contraseña': password,
      'Rol': rol,
      'Sede': sede,
      'Estado': estado,
      'Telefono': telefono
    }
  };

  const btn = document.getElementById('btn-guardar-usuario');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-sm"></span> Guardando...';

  try {
    await enviarAppsScript(payload);
    
    // Mensaje de éxito más claro y amigable
    const mensajeExito = esEdicion 
      ? `✅ El usuario <strong>${nombre}</strong> ha sido <strong>editado</strong> correctamente.` 
      : `✅ El usuario <strong>${nombre}</strong> ha sido <strong>creado</strong> correctamente.`;

    showMsg('msg-usuario', mensajeExito, 'success');

    // Recargar la lista de usuarios
    dataCache[SHEETS.usuarios] = null;
    await cargarUsuarios();

    // Cerrar el formulario después de 1.8 segundos (para que el usuario vea el mensaje)
    setTimeout(() => {
      document.getElementById('form-usuario-card').style.display = 'none';
    }, 1800);

  } catch(err) {
    showMsg('msg-usuario', 'Error al guardar el usuario. Verifica tu conexión.', 'error');
    console.error(err);
  }

  // Restaurar botón
  btn.disabled = false;
  btn.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
      <polyline points="17 21 17 13 7 13 7 21"/>
    </svg> Guardar
  `;
});


// ════════════════════════════════════════════════════════
//  MÓDULO: ESTUDIANTES
// ════════════════════════════════════════════════════════
let estudiantesData  = [];
let editandoEstFila  = -1;   // fila real en Sheets (1-based) o -1 para nuevo

// ── Cargar ────────────────────────────────────────────
async function cargarEstudiantes() {
  document.getElementById('tabla-estudiantes-wrap').innerHTML =
    '<div class="module-loading"><div class="ring-loader"></div><p>Cargando estudiantes...</p></div>';
  try {
    const [dataEst, dataListas] = await Promise.all([
      fetchSheet(SHEETS.estudiantes, true),
      fetchSheet('LISTAS_DESPLEGABLES', true)
    ]);

    if (dataEst.error || !dataEst.values || dataEst.values.length < 2) {
      document.getElementById('tabla-estudiantes-wrap').innerHTML =
        '<div class="empty-state"><p>Sin estudiantes registrados.</p></div>';
      return;
    }

    estudiantesData = dataEst.values;
    const headers   = dataEst.values[0].map(h => h.trim());

    // Poblar filtros de tabla
    const iGrado = headers.findIndex(h => norm(h) === 'GRADO');
    const iSede  = headers.findIndex(h => norm(h) === 'SEDE');
    const rows   = dataEst.values.slice(1);

    const grados = [...new Set(rows.map(r => (r[iGrado]||'').trim()))].filter(Boolean);
    document.getElementById('filtro-grado-est').innerHTML =
      '<option value="">Todos los grados</option>' +
      ordenarGrados(grados).map(g => `<option value="${g}">${g}</option>`).join('');

    const sedes = [...new Set(rows.map(r => (r[iSede]||'').trim()))].filter(Boolean).sort();
    document.getElementById('filtro-sede-est').innerHTML =
      '<option value="">Todas las sedes</option>' +
      sedes.map(s => `<option value="${s}">${s}</option>`).join('');

    // Poblar selects del formulario desde LISTAS_DESPLEGABLES
    if (dataListas.values && dataListas.values.length > 1) {
      const lh   = dataListas.values[0].map(h => h.trim());
      const lrows = dataListas.values.slice(1);

      const iSedeL  = lh.findIndex(h => norm(h) === 'SEDES');
      const iGradoL = lh.findIndex(h => norm(h) === 'GRADOS');

      if (iSedeL >= 0) {
        const opsSede = lrows.map(r => (r[iSedeL]||'').trim()).filter(Boolean);
        document.getElementById('est-sede').innerHTML =
          '<option value="">Seleccionar sede...</option>' +
          opsSede.map(s => `<option value="${s}">${s}</option>`).join('');
      }
      if (iGradoL >= 0) {
        const opsGrado = lrows.map(r => (r[iGradoL]||'').trim()).filter(Boolean);
        document.getElementById('est-grado').innerHTML =
          '<option value="">Seleccionar grado...</option>' +
          opsGrado.map(g => `<option value="${g}">${g}</option>`).join('');
      }
    }

    renderTablaEstudiantes();

  } catch(err) {
    document.getElementById('tabla-estudiantes-wrap').innerHTML =
      '<div class="empty-state"><p>Error al cargar estudiantes.</p></div>';
    console.error(err);
  }
}

// ── Renderizar tabla ──────────────────────────────────
function renderTablaEstudiantes() {
  const filtroGrado  = document.getElementById('filtro-grado-est').value;
  const filtroEstado = document.getElementById('filtro-estado-est').value;
  const filtroSede   = document.getElementById('filtro-sede-est').value;
  const busqueda     = document.getElementById('busqueda-est').value.trim().toLowerCase();

  const headers  = estudiantesData[0].map(h => h.trim());
  const rows     = estudiantesData.slice(1);

  const iNombre   = headers.findIndex(h => norm(h) === 'NOMBRES_APELLIDOS');
  const iGrado    = headers.findIndex(h => norm(h) === 'GRADO');
  const iSede     = headers.findIndex(h => norm(h) === 'SEDE');
  const iEstado   = headers.findIndex(h => norm(h) === 'ESTADO');
  const iDocTipo  = headers.findIndex(h => norm(h) === 'TIPO_DE_DOCUMENTO' || norm(h).includes('TIPO') && norm(h).includes('DOC'));
  const iDocNum   = headers.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
  const iGenero   = headers.findIndex(h => norm(h) === 'GENERO');
  const iGrado2   = iGrado; // alias

  let filtradas = rows.map((row, idx) => ({ row, idx }));

  if (filtroGrado)  filtradas = filtradas.filter(({ row }) => (row[iGrado]||'').trim() === filtroGrado);
  if (filtroSede)   filtradas = filtradas.filter(({ row }) => (row[iSede]||'').trim()  === filtroSede);
  if (filtroEstado) filtradas = filtradas.filter(({ row }) => {
    const e = (row[iEstado]||'').trim().toLowerCase();
    return filtroEstado === 'activo'
      ? (e !== 'inactivo' && e !== 'retirado')
      : (e === 'inactivo' || e === 'retirado');
  });
  if (busqueda) filtradas = filtradas.filter(({ row }) =>
    (row[iNombre]||'').toLowerCase().includes(busqueda) ||
    (row[iDocNum]||'').toLowerCase().includes(busqueda)
  );

  if (!filtradas.length) {
    document.getElementById('tabla-estudiantes-wrap').innerHTML =
      '<div class="empty-state"><p>Sin resultados para los filtros aplicados.</p></div>';
    return;
  }

  document.getElementById('tabla-estudiantes-wrap').innerHTML = `
    <div style="overflow-x:auto;">
      <table class="data-table">
        <thead><tr>
          <th>Documento</th>
          <th>Nombre completo</th>
          <th>Género</th>
          <th>Grado</th>
          <th>Sede</th>
          <th>Estado</th>
          <th>Acciones</th>
        </tr></thead>
        <tbody>
          ${filtradas.map(({ row, idx }) => {
            const estado  = (row[iEstado]||'Activo').trim();
            const activo  = estado.toLowerCase() !== 'inactivo' && estado.toLowerCase() !== 'retirado';
            const filaSheet = idx + 2;
            const docTipo = (row[iDocTipo]||'').trim();
            const docNum  = (row[iDocNum]||'').trim();
            const docStr  = docTipo && docNum ? `${docTipo} ${docNum}` : (docNum || '—');
            return `
            <tr>
              <td style="font-size:12px;color:var(--text-muted);white-space:nowrap;">${docStr}</td>
              <td><strong>${row[iNombre]||'—'}</strong></td>
              <td style="font-size:13px;">${row[iGenero]||'—'}</td>
              <td>${row[iGrado]||'—'}</td>
              <td>${row[iSede]||'—'}</td>
              <td><span class="badge ${activo?'badge-success':'badge-danger'}">${estado}</span></td>
              <td>
                <div style="display:flex;gap:6px;flex-wrap:wrap;">
                  <button class="btn btn-sm" style="background:#1976d2;color:white;border:none;"
                          onclick="editarEstudiante(${filaSheet})">
                    Editar
                  </button>
                  <button class="btn btn-sm"
                          onclick="toggleEstadoEstudiante(${filaSheet}, '${activo ? 'Inactivo' : 'Activo'}', '${(row[iNombre]||'').replace(/'/g,"\\'")}')"
                          style="background:${activo?'#f57c00':'#2e7d32'};color:white;border:none;">
                    ${activo ? 'Inhabilitar' : 'Activar'}
                  </button>
                </div>
              </td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
    <div style="padding:12px 20px;font-size:12px;color:var(--text-muted);">
      Mostrando ${filtradas.length} estudiante${filtradas.length !== 1 ? 's' : ''}
    </div>
  `;
}

// ── Filtros tabla ─────────────────────────────────────
document.getElementById('filtro-grado-est').addEventListener('change', renderTablaEstudiantes);
document.getElementById('filtro-estado-est').addEventListener('change', renderTablaEstudiantes);
document.getElementById('filtro-sede-est').addEventListener('change', renderTablaEstudiantes);
document.getElementById('busqueda-est').addEventListener('input', renderTablaEstudiantes);

// ── Abrir formulario nuevo ────────────────────────────
document.getElementById('btn-nuevo-estudiante').addEventListener('click', () => {
  editandoEstFila = -1;
  document.getElementById('form-estudiante-title').textContent = 'Nuevo estudiante';
  limpiarFormEstudiante();
  document.getElementById('form-estudiante-card').style.display = 'block';
  document.getElementById('form-estudiante-card').scrollIntoView({ behavior: 'smooth' });
});

document.getElementById('btn-cancelar-estudiante').addEventListener('click', () => {
  document.getElementById('form-estudiante-card').style.display = 'none';
});

// ── Limpiar formulario ────────────────────────────────
function limpiarFormEstudiante() {
  ['est-nombres','est-id-doc','est-fecha-nac','est-direccion',
   'est-nombre-madre','est-doc-madre','est-tel-madre',
   'est-nombre-padre','est-doc-padre','est-tel-padre'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  ['est-tipo-doc','est-genero','est-rh','est-sede','est-grado',
   'est-discapacidad','est-vulnerabilidad','est-sisben','est-estrato','est-estado'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.selectedIndex = 0;
  });
  document.getElementById('msg-estudiante').className = 'msg';
}

// ── Editar estudiante ─────────────────────────────────
function editarEstudiante(filaSheet) {
  const rowIdx  = filaSheet - 2;  // índice en estudiantesData.slice(1)
  const row     = estudiantesData.slice(1)[rowIdx];
  if (!row) return;

  const headers = estudiantesData[0].map(h => h.trim());
  const get = (claves) => {
    const i = headers.findIndex(h => claves.some(c => norm(h) === norm(c)));
    return i >= 0 ? (row[i] || '').trim() : '';
  };
  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val;
  };

  editandoEstFila = filaSheet;
  document.getElementById('form-estudiante-title').textContent = 'Editar estudiante';
  limpiarFormEstudiante();

  setVal('est-nombres',      get(['NOMBRES_APELLIDOS']));
  setVal('est-tipo-doc',     get(['TIPO_DE_DOCUMENTO','TIPO DE DOCUMENTO']));
  setVal('est-id-doc',       get(['ID_Estudiante','ID_ESTUDIANTE']));
  setVal('est-genero',       get(['GENERO','GÉNERO']));
  setVal('est-fecha-nac',    (() => {
    const f = get(['FECHA_DE_NACIMIENTO','FECHA DE NACIMIENTO','FECHA NACIMIENTO']);
    if (!f) return '';
    // Convertir dd/mm/yyyy a yyyy-mm-dd para el input date
    const dmy = f.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    return dmy ? `${dmy[3]}-${dmy[2].padStart(2,'0')}-${dmy[1].padStart(2,'0')}` : f;
  })());
  setVal('est-rh',           get(['RH']));
  setVal('est-sede',         get(['SEDE']));
  setVal('est-grado',        get(['GRADO']));
  setVal('est-direccion',    get(['DIRECCION','DIRECCIÓN']));
  setVal('est-discapacidad', get(['DISCAPACIDAD']));
  setVal('est-vulnerabilidad',get(['SITUACION_DE_VULNERABILIDAD','SITUACION DE VULNERABILIDAD']));
  setVal('est-sisben',       get(['CATEGORIA_SISBEN','CATEGORIA SISBEN','CATEGORÍA SISBEN']));
  setVal('est-estrato',      get(['ESTRATO']));
  setVal('est-estado',       get(['ESTADO']) || 'Activo');
  setVal('est-nombre-madre', get(['NOMBRE_MADRE','NOMBRE MADRE']));
  setVal('est-doc-madre',    get(['DOCUMENTO_MADRE','DOCUMENTO MADRE']));
  setVal('est-tel-madre',    get(['TELEFONO_MADRE','TELEFONO MADRE','TELÉFONO MADRE']));
  setVal('est-nombre-padre', get(['NOMBRE_PADRE','NOMBRE PADRE']));
  setVal('est-doc-padre',    get(['DOCUMENTO_PADRE','DOCUMENTO PADRE']));
  setVal('est-tel-padre',    get(['TELEFONO_PADRE','TELEFONO PADRE','TELÉFONO PADRE']));

  document.getElementById('form-estudiante-card').style.display = 'block';
  document.getElementById('form-estudiante-card').scrollIntoView({ behavior: 'smooth' });
}

// ── Toggle estado ─────────────────────────────────────
async function toggleEstadoEstudiante(filaSheet, nuevoEstado, nombre) {
  const texto = nuevoEstado === 'Inactivo' ? 'inhabilitar' : 'activar';
  if (!confirm(`¿Deseas ${texto} a "${nombre}"?`)) return;
  try {
    const resp = await enviarAppsScript({
      tipo:   'toggleEstudiante',
      fila:   filaSheet,
      estado: nuevoEstado
    });
    if (resp.status === 'ok') {
      dataCache[SHEETS.estudiantes] = null;
      await cargarEstudiantes();
    } else {
      alert('Error al cambiar estado: ' + (resp.mensaje || ''));
    }
  } catch(err) {
    alert('Error de conexión.');
    console.error(err);
  }
}

// ── Guardar (nuevo o edición) ─────────────────────────
document.getElementById('btn-guardar-estudiante').addEventListener('click', async () => {
  const nombres  = document.getElementById('est-nombres').value.trim();
  const tipoDoc  = document.getElementById('est-tipo-doc').value;
  const idDoc    = document.getElementById('est-id-doc').value.trim();
  const sede     = document.getElementById('est-sede').value;
  const grado    = document.getElementById('est-grado').value;

  if (!nombres) { showMsg('msg-estudiante', 'El nombre del estudiante es obligatorio.', 'error'); return; }
  if (!tipoDoc) { showMsg('msg-estudiante', 'Selecciona el tipo de documento.', 'error'); return; }
  if (!idDoc)   { showMsg('msg-estudiante', 'El número de documento es obligatorio.', 'error'); return; }
  if (!sede)    { showMsg('msg-estudiante', 'Selecciona la sede.', 'error'); return; }
  if (!grado)   { showMsg('msg-estudiante', 'Selecciona el grado.', 'error'); return; }

  // Fecha: de yyyy-mm-dd a dd/mm/yyyy
  const fechaRaw = document.getElementById('est-fecha-nac').value;
  let fechaFmt = '';
  if (fechaRaw) {
    const [y,m,d] = fechaRaw.split('-');
    fechaFmt = `${d}/${m}/${y}`;
  }

  const btn = document.getElementById('btn-guardar-estudiante');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-sm"></span> Guardando...';

  const datos = {
    'NOMBRES_APELLIDOS':           nombres,
    'SEDE':                        sede,
    'GENERO':                      document.getElementById('est-genero').value,
    'GRADO':                       grado,
    'FECHA DE NACIMIENTO':         fechaFmt,
    'TIPO DE DOCUMENTO':           tipoDoc,
    'ID_Estudiante':               idDoc,
    'DIRECCION':                   document.getElementById('est-direccion').value.trim(),
    'RH':                          document.getElementById('est-rh').value,
    'DISCAPACIDAD':                document.getElementById('est-discapacidad').value,
    'SITUACION DE VULNERABILIDAD': document.getElementById('est-vulnerabilidad').value,
    'CATEGORIA SISBEN':            document.getElementById('est-sisben').value,
    'ESTRATO':                     document.getElementById('est-estrato').value,
    'ESTADO':                      document.getElementById('est-estado').value,
    'NOMBRE MADRE':                document.getElementById('est-nombre-madre').value.trim(),
    'DOCUMENTO MADRE':             document.getElementById('est-doc-madre').value.trim(),
    'TELEFONO MADRE':              document.getElementById('est-tel-madre').value.trim(),
    'NOMBRE PADRE':                document.getElementById('est-nombre-padre').value.trim(),
    'DOCUMENTO PADRE':             document.getElementById('est-doc-padre').value.trim(),
    'TELEFONO PADRE':              document.getElementById('est-tel-padre').value.trim(),
  };

  try {
    const resp = await enviarAppsScript({
      tipo: 'guardarEstudiante',
      fila: editandoEstFila > 0 ? editandoEstFila : -1,
      datos
    });

    if (resp.status === 'ok') {
      const accion = editandoEstFila > 0 ? 'actualizado' : 'registrado';
      showMsg('msg-estudiante', `✅ Estudiante <strong>${nombres}</strong> ${accion} correctamente.`, 'success');
      dataCache[SHEETS.estudiantes] = null;
      await cargarEstudiantes();
      setTimeout(() => {
        document.getElementById('form-estudiante-card').style.display = 'none';
      }, 1800);
    } else {
      showMsg('msg-estudiante', `Error: ${resp.mensaje || 'Intenta de nuevo.'}`, 'error');
    }
  } catch(err) {
    showMsg('msg-estudiante', 'Error de conexión. Verifica tu red.', 'error');
    console.error(err);
  }

  btn.disabled = false;
  btn.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
      <polyline points="17 21 17 13 7 13 7 21"/>
    </svg> Guardar estudiante`;
});

// ════════════════════════════════════════════════════════
//  MÓDULO: PERIODOS - Versión Profesional + Calendario Visual
// ════════════════════════════════════════════════════════
let periodosData = [];

async function cargarPeriodos() {
  const wrap = document.getElementById('tabla-periodos-wrap');
  wrap.innerHTML = '<div class="module-loading"><div class="ring-loader"></div><p>Cargando periodos...</p></div>';

  try {
    const data = await fetchSheet(SHEETS.periodos, true);
    if (data.error || !data.values || data.values.length < 2) {
      wrap.innerHTML = '<div class="empty-state"><p>No hay datos de periodos configurados.</p></div>';
      return;
    }
    
    periodosData = data.values;
    renderPeriodosMejorado();   // ← Esta es la función clave
    
  } catch(err) {
    wrap.innerHTML = '<div class="empty-state"><p>Error al cargar los periodos.</p></div>';
    console.error('Error cargando periodos:', err);
  }
}

function renderPeriodosMejorado() {
  const headers = periodosData[0].map(h => h.trim());
  const rows = periodosData.slice(1);
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const iNum = headers.findIndex(h => norm(h).includes('PERIODO'));
  const iIni = headers.findIndex(h => norm(h).includes('INICIO'));
  const iFin = headers.findIndex(h => norm(h).includes('FINAL') && !norm(h).includes('CIERRE'));
  const iCie = headers.findIndex(h => norm(h).includes('CIERRE'));

  document.getElementById('btn-guardar-periodos').addEventListener('click', async () => {
  const inputs  = document.querySelectorAll('.periodo-row-input');
  const cambios = [];

  inputs.forEach(input => {
    cambios.push({
      idx:    input.dataset.idx,
      campo:  input.dataset.campo,
      valor:  input.value.trim()
    });
  });

  const btn = document.getElementById('btn-guardar-periodos');
  btn.disabled = true;
  btn.textContent = 'Guardando...';

  try {
    await enviarAppsScript({ tipo: 'guardarPeriodos', cambios });
    showMsg('msg-periodos', '✅ Periodos guardados correctamente.', 'success');
    dataCache[SHEETS.periodos] = null;
    await cargarPeriodos();
  } catch(err) {
    showMsg('msg-periodos', 'Error al guardar los periodos.', 'error');
  }

  btn.disabled = false;
  btn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/></svg> Guardar Cambios`;
});

  // ==================== CALENDARIO VISUAL ====================
  let calendarHTML = `<div class="periodo-calendar">`;

  rows.forEach((row, idx) => {
    const num = String(row[iNum] || idx + 1).trim();
    const ini = parseFecha(row[iIni]);
    const fin = parseFecha(row[iFin]);
    const cie = parseFecha(row[iCie]);

    let estadoClass = 'estado-futuro';
    let estadoTexto = 'Futuro';
    let icono = '⏳';

    if (ini && fin) {
      if (hoy >= ini && hoy <= fin) {
        if (cie && hoy > cie) {
          estadoClass = 'estado-cerrado';
          estadoTexto = 'En curso (Cerrado)';
          icono = '🔒';
        } else {
          estadoClass = 'estado-abierto';
          estadoTexto = 'En curso (Abierto)';
          icono = '▶️';
        }
      } else if (hoy > fin) {
        estadoClass = 'estado-finalizado';
        estadoTexto = 'Finalizado';
        icono = '✅';
      }
    }

    calendarHTML += `
      <div>
        <strong>Periodo ${num}</strong>
        <div class="periodo-fecha">
          ${formatFecha(ini) || '—'} → ${formatFecha(fin) || '—'}
        </div>
        <div style="margin-top:8px; font-size:13px;">
          <span style="margin-right:6px;">${icono}</span>
          <span class="${estadoClass}">${estadoTexto}</span>
        </div>
      </div>`;
  });

  calendarHTML += `</div>`;

  // Insertar calendario
  const calendarDiv = document.getElementById('periodo-calendar');
  if (calendarDiv) {
    calendarDiv.innerHTML = calendarHTML;
  }

  // ==================== TABLA ====================
  let tablaHTML = `
    <div style="overflow-x:auto;">
      <table class="data-table">
        <thead>
          <tr>
            <th>Periodo</th>
            <th>Fecha Inicio</th>
            <th>Fecha Final</th>
            <th>Cierre del Sistema</th>
            <th>Estado Actual</th>
          </tr>
        </thead>
        <tbody>`;

  rows.forEach((row, idx) => {
    const ini = parseFecha(row[iIni]);
    const fin = parseFecha(row[iFin]);
    const cie = parseFecha(row[iCie]);

    let badgeHTML = `<span class="badge badge-gray">Futuro</span>`;
    if (ini && fin) {
      if (hoy >= ini && hoy <= fin) {
        badgeHTML = cie && hoy > cie 
          ? `<span class="badge badge-warning">En curso • Cerrado</span>`
          : `<span class="badge badge-success">En curso • Abierto</span>`;
      } else if (hoy > fin) {
        badgeHTML = `<span class="badge badge-gray">Finalizado</span>`;
      }
    }

    tablaHTML += `
      <tr>
        <td><strong>Periodo ${row[iNum] || idx+1}</strong></td>
        <td><input type="text" class="periodo-row-input" data-campo="inicio" data-idx="${idx}" value="${formatFecha(ini)||''}" placeholder="DD/MM/YYYY"/></td>
        <td><input type="text" class="periodo-row-input" data-campo="final"  data-idx="${idx}" value="${formatFecha(fin)||''}" placeholder="DD/MM/YYYY"/></td>
        <td><input type="text" class="periodo-row-input" data-campo="cierre" data-idx="${idx}" value="${formatFecha(cie)||''}" placeholder="DD/MM/YYYY"/></td>
        <td>${badgeHTML}</td>
      </tr>`;
  });

  tablaHTML += `</tbody></table></div>`;

  document.getElementById('tabla-periodos-wrap').innerHTML = tablaHTML;
}

// ════════════════════════════════════════════════════════
//  MÓDULO: CONSOLIDADO
// ════════════════════════════════════════════════════════
async function cargarConsolidado() {
  try {
    const data = await fetchSheet(SHEETS.consolidado, true);
    if (!data.error && data.values && data.values.length > 1) {
      const headers = data.values[0].map(h => (h||'').trim());
      const iGrado  = headers.findIndex(h => norm(h) === 'GRADO');
      const grados  = [...new Set(data.values.slice(1).map(r => (r[iGrado]||'').trim()).filter(Boolean))];
      const selGrado = document.getElementById('filtro-grado-consol');
      selGrado.innerHTML =
        '<option value="">Selecciona un grado</option>' +
        '<option value="todos">Todos los grados</option>' +
        ordenarGrados(grados).map(g => `<option value="${g}">${g}</option>`).join('');

      const iSede = headers.findIndex(h => norm(h) === 'SEDE');
      const sedes = [...new Set(data.values.slice(1).map(r => (r[iSede]||'').trim()).filter(Boolean))].sort();
      const selSede = document.getElementById('filtro-sede-consol');
      selSede.innerHTML =
        '<option value="">Selecciona una sede</option>' +
        '<option value="todas">Todas las sedes</option>' +
        sedes.map(s => `<option value="${s}">${s}</option>`).join('');

      // Event listeners para los filtros
      selGrado.addEventListener('change', renderTablaConsolidado);
      selSede.addEventListener('change', renderTablaConsolidado);
      const busq = document.getElementById('busqueda-consol');
      if (busq) busq.addEventListener('input', renderTablaConsolidado);
    }

    // Mostrar estado inicial vacío hasta que el usuario filtre
    document.getElementById('tabla-consolidado-wrap').innerHTML =
      '<div class="empty-state"><p>Selecciona un grado o sede para ver el consolidado.</p></div>';

  } catch(err) {
    console.error('Error cargando consolidado:', err);
  }
}

async function renderTablaConsolidado() {
  const gradoFiltro = document.getElementById('filtro-grado-consol').value;
  const sedeFiltro  = document.getElementById('filtro-sede-consol').value;
  const busqueda    = (document.getElementById('busqueda-consol')?.value || '').toLowerCase();

  if (!gradoFiltro && !sedeFiltro && !busqueda) {
    document.getElementById('tabla-consolidado-wrap').innerHTML =
      '<div class="empty-state"><p>Selecciona un grado o sede para ver el consolidado.</p></div>';
    return;
  }

  document.getElementById('tabla-consolidado-wrap').innerHTML =
    '<div class="module-loading"><div class="ring-loader"></div><p>Cargando consolidado...</p></div>';

  try {
    const data = await fetchSheet(SHEETS.consolidado, true);
    if (data.error || !data.values || data.values.length < 2) {
      document.getElementById('tabla-consolidado-wrap').innerHTML =
        '<div class="empty-state"><p>No hay datos en el consolidado.</p></div>';
      return;
    }

    const headers = data.values[0].map(h => (h||'').trim());
    const rows    = data.values.slice(1);

    const iGradoC = headers.findIndex(h => norm(h) === 'GRADO');
    const iSedeC  = headers.findIndex(h => norm(h) === 'SEDE');
    const iNombre = headers.findIndex(h => norm(h).includes('NOMBRE'));

    let filtradas = rows;
    if (gradoFiltro && gradoFiltro !== 'todos')
      filtradas = filtradas.filter(r => (r[iGradoC]||'').trim().toLowerCase() === gradoFiltro.toLowerCase());
    if (sedeFiltro && sedeFiltro !== 'todas')
      filtradas = filtradas.filter(r => (r[iSedeC]||'').trim().toLowerCase() === sedeFiltro.toLowerCase());
    if (busqueda)
      filtradas = filtradas.filter(r => (r[iNombre]||'').toLowerCase().includes(busqueda));

    if (!filtradas.length) {
      document.getElementById('tabla-consolidado-wrap').innerHTML =
        '<div class="empty-state"><p>No hay registros que coincidan con los filtros.</p></div>';
      return;
    }

    const colsAsig = headers
      .map((h, i) => ({ h, i }))
      .filter(c => c.h.toUpperCase().includes('_FINAL') && !c.h.toUpperCase().includes('DESEMPE'));

// ════════════════════════════════════════════════════════
    // MOSTRAR TODAS LAS COLUMNAS (como estaba antes)
    // ════════════════════════════════════════════════════════
    document.getElementById('tabla-consolidado-wrap').innerHTML = `
      <div style="overflow-x:auto;">
        <table class="data-table" style="min-width:1400px;">
          <thead><tr>
            ${headers.map(h => `<th>${h}</th>`).join('')}
          </tr></thead>
          <tbody>
            ${filtradas.map(r => `
              <tr>
                ${r.map(celda => `<td>${celda || ''}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  } catch(err) {
    console.error('Error renderizando consolidado:', err);
    document.getElementById('tabla-consolidado-wrap').innerHTML =
      '<div class="empty-state"><p>Error al cargar los datos.</p></div>';
  }
}

// Botón Cargar (por si existe en el HTML)
const btnCargarConsol = document.getElementById('btn-cargar-consolidado');
if (btnCargarConsol) btnCargarConsol.addEventListener('click', renderTablaConsolidado);

// ════════════════════════════════════════════════════════
//  MÓDULO: BOLETINES - Versión mínima y estable
// ════════════════════════════════════════════════════════

let estudiantesBoletin = [];
let headersEstudiantesBoletin = []; // headers sincronizados con estudiantesBoletin

// Cargar filtros
async function cargarFiltrosBoletines() {
  try {
    const data = await fetchSheet(SHEETS.estudiantes);
    if (data.error || !data.values || data.values.length < 2) return;

    const headers = data.values[0].map(h => h.trim());
    const rows = data.values.slice(1);

    const iSede = headers.findIndex(h => norm(h) === 'SEDE');
    const iGrado = headers.findIndex(h => norm(h) === 'GRADO');

    // Sedes
    const sedes = [...new Set(rows.map(r => (r[iSede] || '').trim()).filter(Boolean))].sort();
    document.getElementById('filtro-sede-boletin').innerHTML = `
      <option value="">Todas las sedes</option>
      ${sedes.map(s => `<option value="${s}">${s}</option>`).join('')}
    `;

    // Grados - Predeterminado "Todos los grados"
    const grados = [...new Set(rows.map(r => (r[iGrado] || '').trim()).filter(Boolean))];
    const selGrado = document.getElementById('filtro-grado-boletin');
    let htmlGrado = '<option value="todos" selected>Todos los grados</option>';
    ordenarGrados(grados).forEach(g => {
      htmlGrado += `<option value="${g}">${g}</option>`;
    });
    selGrado.innerHTML = htmlGrado;

    // Periodo
    document.getElementById('filtro-periodo-boletin').innerHTML = `
      <option value="">Selecciona periodo</option>
      <option value="1">Periodo 1</option>
      <option value="2">Periodo 2</option>
      <option value="3">Periodo 3</option>
      <option value="4">Periodo 4</option>
      <option value="final">Boletín Final Anual</option>
    `;

    // Eventos simples (sin debounce complicado)
    document.getElementById('filtro-sede-boletin').addEventListener('change', () => {
      const sede = document.getElementById('filtro-sede-boletin').value;
      const data = dataCache[SHEETS.estudiantes];
      if (data && data.values) {
        const headers = data.values[0].map(h => h.trim());
        const rows = data.values.slice(1);
        const iSede  = headers.findIndex(h => norm(h) === 'SEDE');
        const iGrado = headers.findIndex(h => norm(h) === 'GRADO');
        const filtradas = sede ? rows.filter(r => (r[iSede] || '').trim() === sede) : rows;
        const grados = [...new Set(filtradas.map(r => (r[iGrado] || '').trim()).filter(Boolean))];
        const selGrado = document.getElementById('filtro-grado-boletin');
        let htmlGrado = '<option value="todos" selected>Todos los grados</option>';
        ordenarGrados(grados).forEach(g => { htmlGrado += `<option value="${g}">${g}</option>`; });
        selGrado.innerHTML = htmlGrado;
      }
      cargarEstudiantesBoletinesAutomatico();
    });
    document.getElementById('filtro-grado-boletin').addEventListener('change', cargarEstudiantesBoletinesAutomatico);
    document.getElementById('filtro-periodo-boletin').addEventListener('change', cargarEstudiantesBoletinesAutomatico);

    const selectGrado = document.getElementById('filtro-grado-boletin');
if (selectGrado) {
    selectGrado.disabled = false;
    selectGrado.style.pointerEvents = 'auto';
    selectGrado.style.opacity = '1';
    selectGrado.style.backgroundColor = 'white';
    console.log('✅ Selector de grados habilitado. Opciones:', selectGrado.options.length);
}


    // Carga inicial
    setTimeout(cargarEstudiantesBoletinesAutomatico, 600);

  } catch(err) {
    console.error('Error cargando filtros:', err);
  }
}

// Carga de estudiantes automática
async function cargarEstudiantesBoletinesAutomatico() {
  const sede = document.getElementById('filtro-sede-boletin').value;
  const grado = document.getElementById('filtro-grado-boletin').value;
  const periodo = document.getElementById('filtro-periodo-boletin').value;

  if (!periodo) {
    document.getElementById('tabla-estudiantes-boletin-wrap').innerHTML = 
      '<div class="empty-state"><p>Selecciona un periodo.</p></div>';
    document.getElementById('lista-estudiantes-boletin').style.display = 'none';
    return;
  }

  const wrap = document.getElementById('tabla-estudiantes-boletin-wrap');
  wrap.innerHTML = '<div class="module-loading"><div class="ring-loader"></div><p>Cargando...</p></div>';

  try {
    const data = await fetchSheet(SHEETS.estudiantes, true);
    if (data.error || !data.values) throw new Error("Error cargando estudiantes");

    const headers = data.values[0].map(h => h.trim());
    const rows = data.values.slice(1);

    const iNombre = headers.findIndex(h => norm(h).includes('NOMBRE') || norm(h).includes('APELLIDO'));
    const iGradoIdx = headers.findIndex(h => norm(h) === 'GRADO');
    const iSedeIdx = headers.findIndex(h => norm(h) === 'SEDE');
    const iDocumento = headers.findIndex(h => norm(h) === 'ID_ESTUDIANTE' || norm(h).includes('DOCUMENTO'));

    let filtrados = rows;
    if (sede) filtrados = filtrados.filter(r => (r[iSedeIdx] || '').trim() === sede);
    if (grado && grado !== "todos") filtrados = filtrados.filter(r => (r[iGradoIdx] || '').trim() === grado);

    estudiantesBoletin = filtrados;
    headersEstudiantesBoletin = headers; // ← guardar headers sincronizados

    if (filtrados.length === 0) {
      wrap.innerHTML = '<div class="empty-state"><p>No hay estudiantes con estos filtros.</p></div>';
      document.getElementById('lista-estudiantes-boletin').style.display = 'none';
      return;
    }

    let html = `<table class="data-table">
      <thead><tr>
        <th>No Documento</th>
        <th>Nombre Completo</th>
        <th>Sede</th>
        <th>Grado</th>
        <th>Acciones</th>
      </tr></thead><tbody>`;

    filtrados.forEach((row, idx) => {
      const nombre = row[iNombre] || 'Sin nombre';
      const doc = row[iDocumento] && row[iDocumento] !== '' ? row[iDocumento] : '—';
      html += `<tr>
        <td><strong>${doc}</strong></td>
        <td>${nombre}</td>
        <td>${row[iSedeIdx] || ''}</td>
        <td>${row[iGradoIdx] || ''}</td>
        <td><button class="btn btn-primary btn-sm" onclick="verPreviaBoletin(${idx})">Ver Boletín</button></td>
      </tr>`;
    });

    html += `</tbody></table>`;
    wrap.innerHTML = html;
    document.getElementById('lista-estudiantes-boletin').style.display = 'block';

  } catch(err) {
    wrap.innerHTML = '<div class="empty-state"><p>Error al cargar.</p></div>';
    console.error(err);
  }
}

// Ocultar botón antiguo
if (document.getElementById('btn-cargar-boletines')) {
  document.getElementById('btn-cargar-boletines').style.display = 'none';
}

// =============================================
//  FUNCIÓN COMPLETA: verPreviaBoletin
// =============================================
async function verPreviaBoletin(idx, silencioso = false) {
  const estudiante = estudiantesBoletin[idx];
  if (!estudiante) return;

  const periodoSeleccionado = document.getElementById('filtro-periodo-boletin').value;
  const esFinalAnual = periodoSeleccionado === "final";
  const numPeriodo = esFinalAnual ? 4 : parseInt(periodoSeleccionado);

  // Usar los headers guardados en el momento en que se cargó la lista de estudiantes
  // (evita desincronización de índices por caché entre fetchSheet y el array estudiantesBoletin)
  const headers = headersEstudiantesBoletin.length > 0
    ? headersEstudiantesBoletin
    : (await fetchSheet(SHEETS.estudiantes, true)).values[0].map(h => h.trim());

  const iNombreB   = headers.findIndex(h => norm(h).includes('NOMBRE'));
  const iDocumento = headers.findIndex(h => norm(h) === 'ID_ESTUDIANTE' || norm(h).includes('DOCUMENTO'));
  const iSedeB     = headers.findIndex(h => norm(h) === 'SEDE');
  const iGradoB    = headers.findIndex(h => norm(h) === 'GRADO');

  const nombreEst  = (iNombreB   >= 0 ? estudiante[iNombreB]   : '') || 'Sin nombre';
  const documento  = (iDocumento >= 0 ? estudiante[iDocumento] : '') || '—';
  console.log('documento:', documento, '| iDocumento:', iDocumento, '| headers:', headers);
  const sedeEst    = (iSedeB     >= 0 ? estudiante[iSedeB]     : '') || 'Sede Principal';
  const gradoEst   = (iGradoB    >= 0 ? estudiante[iGradoB]    : '') || '';

  const nivelEst = getNivelFromGrado(gradoEst);

  showMsg('msg-boletines', `Cargando boletín de ${nombreEst}...`, 'info');

  try {
    // 1. Configuración_boletines
    const configData = await fetchSheet(SHEETS.configBoletines, true);
    const configHeaders = configData.values[0].map(h => h.trim());
    const configRows = configData.values.slice(1);

    const iNivelConfig = configHeaders.findIndex(h => norm(h) === 'NIVEL');
    const iAsigConfig  = configHeaders.findIndex(h => norm(h) === 'ASIGNATURA');
    const iAreaConfig  = configHeaders.findIndex(h => norm(h) === 'AREA');
    const iPorcConfig  = configHeaders.findIndex(h => norm(h) === 'PORCENTAJE');
    const iOrdenConfig = configHeaders.findIndex(h => norm(h) === 'ORDEN');

    // 2. Tipo de Sede
    const sedesData = await fetchSheet('LISTAS_DESPLEGABLES', true);
    let esSedeUnitaria = false;
    if (sedesData.values && sedesData.values.length > 1) {
      const sedesHeaders = sedesData.values[0].map(h => h.trim());
      const sedesRows = sedesData.values.slice(1);
      const iSedeCol = sedesHeaders.findIndex(h => norm(h) === 'SEDES');
      const iTipoCol = sedesHeaders.findIndex(h => norm(h) === 'SEDE UNITARIA' || norm(h) === 'TIPO SEDE');

      const sedeRow = sedesRows.find(row => norm(row[iSedeCol] || '') === norm(sedeEst));
      if (sedeRow && iTipoCol !== -1) {
        esSedeUnitaria = (sedeRow[iTipoCol] || '').toString().trim().toUpperCase() === 'SI';
      }
    }

    // 3. Director de Grado
    const estData = await fetchSheet(SHEETS.estudiantes, true);
    const estHeaders = estData.values[0].map(h => h.trim());
    const estRows = estData.values.slice(1);
    const iDocEst = estHeaders.findIndex(h => norm(h) === 'ID_ESTUDIANTE' || norm(h).includes('DOCUMENTO'));
    const iDirector = estHeaders.findIndex(h => norm(h).includes('COMPORTAMIENTO'));
    let directorGrado = 'Sin asignar';
    const filaEst = estRows.find(row => String(row[iDocEst] || '').trim() === documento);
    if (filaEst && iDirector !== -1) directorGrado = filaEst[iDirector] || 'Sin asignar';

    // 4b. Indicadores de logro por asignatura
    const indData = await fetchSheet('Indicadores', true);
    const mapaIndicadores = {}; // { "ASIGNATURA": "texto del indicador" }
    if (indData.values && indData.values.length > 1) {
      const indHeaders = indData.values[0].map(h => h.trim());
      const indRows    = indData.values.slice(1);
      const iIndPer  = indHeaders.findIndex(h => norm(h) === 'PERIODO');
      const iIndGra  = indHeaders.findIndex(h => norm(h) === 'GRADO');
      const iIndAsig = indHeaders.findIndex(h => norm(h) === 'ASIGNATURA');
      const iIndTxt  = indHeaders.findIndex(h => norm(h) === 'INDICADOR');
      const iIndDoc  = indHeaders.findIndex(h => norm(h) === 'DOCENTE');
      indRows.forEach(row => {
        if (String(row[iIndPer]  || '').trim() !== String(periodoSeleccionado).trim()) return;
        const gradoFila = String(row[iIndGra] || '').trim();
        const gradoComp = String(gradoEst || '').trim();
        if (!gradoFila || !gradoComp) return;
        if (norm(gradoFila) !== norm(gradoComp)) return;
        const asig = (row[iIndAsig] || '').trim();
        const txt  = (row[iIndTxt]  || '').trim();
        if (!asig || !txt) return;
        // Si hay docente en el indicador y en la fila del estudiante, filtrar por docente
        if (iIndDoc >= 0 && filaEst) {
          const docenteIndicador = String(row[iIndDoc] || '').trim();
          const iColAsig = estHeaders.findIndex(h => norm(h) === norm(asig));
          const docenteEstudiante = iColAsig >= 0 ? String(filaEst[iColAsig] || '').trim() : '';
          if (docenteIndicador && docenteEstudiante && docenteIndicador !== docenteEstudiante) return;
        }
        mapaIndicadores[norm(asig)] = txt;
      });
    }

    // 4. Observaciones
    const obsData = await fetchSheet('Observaciones_Boletines', true);
    let observacion = '';
    if (obsData.values && obsData.values.length > 1) {
      const obsHeaders = obsData.values[0].map(h => h.trim());
      const obsRows    = obsData.values.slice(1);
      const iDocObs = obsHeaders.findIndex(h => norm(h) === 'DOCUMENTO');
      const iPerObs = obsHeaders.findIndex(h => norm(h) === 'PERIODO');
      const iObs    = obsHeaders.findIndex(h => norm(h) === 'OBSERVACION');
      const obsEncontrada = obsRows.find(row =>
        String(row[iDocObs] || '').trim() === documento &&
        String(row[iPerObs] || '').trim() === periodoSeleccionado
      );
      if (obsEncontrada && iObs !== -1) observacion = obsEncontrada[iObs] || '';
    }

    // ── seqEst: número secuencial del estudiante (clave para consolidados) ──
    const iSeqEst = estHeaders.findIndex(h => norm(h) === 'NO_IDENTIFICADOR');
    const seqEst  = (iSeqEst >= 0 && filaEst)
      ? String(filaEst[iSeqEst] || '').trim()
      : '';

    // 5. Consolidado calificaciones
    const consData    = await fetchSheet(SHEETS.consolidado, true);
    const consHeaders = consData.values ? consData.values[0].map(h => h.trim()) : [];
    const consRows    = consData.values ? consData.values.slice(1) : [];
    const iSeqCons    = consHeaders.findIndex(h => norm(h) === 'NO_IDENTIFICADOR');
    let filaCons = null;
    if (consData.values && iSeqCons >= 0 && seqEst) {
      filaCons = consRows.find(row => String(row[iSeqCons] || '').trim() === seqEst);
    }
    if (!filaCons) {
      const iDocCons = consHeaders.findIndex(h =>
        norm(h).includes('DOCUMENTO') || norm(h) === 'ID_ESTUDIANTE'
      );
      if (iDocCons >= 0)
        filaCons = consRows.find(row => String(row[iDocCons] || '').trim() === documento);
    }

    // 5b. Calificaciones — mapa completo por asignatura y periodo
    // mapaCalif[asignatura][periodo] = { promedio, recuperacion, definitiva }
    const califData    = await fetchSheet('Calificaciones', true);
    const califHeaders = califData.values ? califData.values[0].map(h => h.trim()) : [];
    const califRows    = califData.values ? califData.values.slice(1) : [];
    const iCalifDoc    = califHeaders.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
    const iCalifPer    = califHeaders.findIndex(h => norm(h) === 'PERIODO');
    const iCalifAsig   = califHeaders.findIndex(h => norm(h) === 'ASIGNATURA');
    const iCalifRec    = califHeaders.findIndex(h => norm(h) === 'RECUPERACION');
    const iCalifProm   = califHeaders.findIndex(h => norm(h) === 'PROMEDIO DEL PERIODO' || norm(h) === 'PROMEDIO_DEL_PERIODO');
    const mapaCalif    = {}; // { asig: { 1: {promedio, recuperacion, definitiva}, 2: {...} } }
    if (califData.values && califData.values.length > 1) {
      califRows.forEach(row => {
        if (String(row[iCalifDoc] || '').trim() !== documento) return;
        const asig = (row[iCalifAsig] || '').trim();
        const per  = String(row[iCalifPer] || '').trim();
        if (!asig || !per) return;
        const prom = parseFloat(String(row[iCalifProm] || '').replace(',','.')) || 0;
        const rec  = parseFloat(String(row[iCalifRec]  || '').replace(',','.')) || 0;
        const def  = rec > 0 ? Math.max(prom, rec) : prom;
        const asigKey = norm(asig); // clave normalizada para comparación robusta
        if (!mapaCalif[asigKey]) mapaCalif[asigKey] = {};
        mapaCalif[asigKey][per] = {
          promedio:     prom > 0 ? prom.toFixed(1) : '—',
          recuperacion: rec  > 0 ? rec.toFixed(1)  : '—',
          definitiva:   def  > 0 ? def.toFixed(1)  : '—'
        };
      });
    }
    // Compatibilidad: mapaRecuperacion para el periodo seleccionado (usado en lógica de área)
    const mapaRecuperacion = {};
    Object.entries(mapaCalif).forEach(([asig, periodos]) => {
      const p = periodos[String(periodoSeleccionado)];
      if (p && p.recuperacion !== '—') mapaRecuperacion[asig] = p.recuperacion;
      if (p && p.recuperacion !== '—') mapaRecuperacion[norm(asig)] = p.recuperacion;
    });

    // 5c. CONSOLIDADO_ASISTENCIA (fallas por asignatura y periodo)
    const asistData    = await fetchSheet('CONSOLIDADO_ASISTENCIA', true);
    const asistHeaders = asistData.values ? asistData.values[0].map(h => h.trim()) : [];
    const asistRows    = asistData.values ? asistData.values.slice(1) : [];
    const iSeqAsist    = asistHeaders.findIndex(h => norm(h) === 'NO_IDENTIFICADOR');
    let filaAsist = null;
    if (asistData.values && iSeqAsist >= 0 && seqEst) {
      filaAsist = asistRows.find(row => String(row[iSeqAsist] || '').trim() === seqEst);
    }

    // 6. Promedio general del periodo — desde CONSOLIDADO columna PROMEDIO_GENERAL_Pn
    let promedioGeneral = '—';
    if (filaCons) {
      const colPG = 'PROMEDIO_GENERAL_P' + (esFinalAnual ? 4 : numPeriodo);
      const iPG   = consHeaders.findIndex(h => norm(h) === norm(colPG));
      if (iPG >= 0) {
        const raw = String(filaCons[iPG] || '').replace(',', '.');
        const val = parseFloat(raw);
        if (!isNaN(val) && val > 0) promedioGeneral = val.toFixed(1);
      }
    }

    // 7. Ranking — basado en PROMEDIO_GENERAL_Pn del Consolidado
    let ranking = '—';
    if (filaCons) {
      const colPGRank = 'PROMEDIO_GENERAL_P' + (esFinalAnual ? 4 : numPeriodo);
      const iPGRank   = consHeaders.findIndex(h => norm(h) === norm(colPGRank));
      // Fallback: buscar PROMEDIO DEL PERIODO o PROMEDIO
      const iPromedioR = iPGRank >= 0 ? iPGRank : consHeaders.findIndex(h =>
        norm(h) === 'PROMEDIO DEL PERIODO' || norm(h) === 'PROMEDIO_DEL_PERIODO' || norm(h) === 'PROMEDIO'
      );
      if (iPromedioR !== -1) {
        const iGradoCons    = consHeaders.findIndex(h => norm(h) === 'GRADO');
        const iNomCons      = consHeaders.findIndex(h => norm(h).includes('NOMBRE'));
        const todosDelGrado = consRows.filter(row =>
          row[iGradoCons] && norm(row[iGradoCons]) === norm(gradoEst)
        );

        // Crear lista con promedio y nombre para desempate alfabético
        const listaRanking = todosDelGrado
          .map(row => ({
            promedio: parseFloat(String(row[iPromedioR]||'').replace(',','.')) || 0,
            nombre:   iNomCons >= 0 ? String(row[iNomCons]||'').trim() : ''
          }))
          .filter(e => e.promedio > 0)
          .sort((a, b) => {
            if (b.promedio !== a.promedio) return b.promedio - a.promedio;
            return a.nombre.localeCompare(b.nombre); // desempate alfabético
          });

        const miPromedio = parseFloat(String(filaCons[iPromedioR]||'').replace(',','.')) || 0;
        if (miPromedio > 0 && listaRanking.length > 0) {
          // Posición: buscar la posición del estudiante actual
          const miNombre = iNomCons >= 0 ? String(filaCons[iNomCons]||'').trim() : '';
          const posicion = listaRanking.findIndex(e =>
            e.promedio === miPromedio && e.nombre === miNombre
          ) + 1;
          ranking = posicion > 0 ? `${posicion}/${listaRanking.length}` : '—';
        }
      }
    }

    // Mapa normalizado: norm(nombre_asignatura) → codigo_consolidado
    const MAPA_BOL_RAW = {
      "Matemáticas":                                "MATEMATICAS",
      "Lengua Castellana":                          "LENGUAJE",
      "Ciencias Naturales y Educación Ambiental":   "CCNN",
      "Ciencias Sociales":                          "CCSS",
      "Cátedra para la paz":                        "CATEDRA",
      "Educación Artística y Cultural":             "ARTISTICA",
      "Educación Ética y en valores humanos":       "ETICA",
      "Educación Física, recreación y deportes":    "EDUFISICA",
      "Educación Religiosa":                        "RELIGION",
      "Inglés":                                     "INGLES",
      "Geometría":                                  "GEOMETRIA",
      "Tecnología e Informática":                   "TECNOLOGIA",
      "Emprendimiento":                             "EMPRENDIMIENTO",
      "Estadística":                                "ESTADISTICA",
      "Física":                                     "FISICA",
      "Química":                                    "QUIMICA",
      "Álgebra":                                    "ALGEBRA",
      "Filosofía":                                  "FILOSOFIA",
      "Ciencias Económicas y Políticas":            "ECONOMICAS",      
      "Centro de Interés":                          "CI",
      "Comportamiento":                             "COMPORTAMIENTO"
    };
    // Construir mapa con claves normalizadas para comparación insensible
    const MAPA_BOL = {};
    Object.entries(MAPA_BOL_RAW).forEach(([k, v]) => { MAPA_BOL[norm(k)] = v; });

    // Busca el índice de la columna para una asignatura+periodo en unos headers dados
    function findColIdx(headers, asignatura, periodo) {
      const sufijo = '_P' + periodo;
      // 1. Buscar por MAPA_BOL (código corto, ej: MATEMATICAS_P1)
      const codigo = MAPA_BOL[asignatura];
      if (codigo) {
        const i = headers.findIndex(h => norm(h) === norm(codigo + sufijo));
        if (i >= 0) return i;
      }
      // 2. Fallback: buscar por nombre normalizado de la asignatura (ej: MATEMATICAS_P1 o LENGUA CASTELLANA_P1)
      const normAsig = norm(asignatura).replace(/\s+/g, '');
      const i2 = headers.findIndex(h => {
        const hn = norm(h).replace(/\s+/g, '');
        return hn === normAsig + 'P' + periodo || hn === normAsig + '_P' + periodo;
      });
      if (i2 >= 0) return i2;
      // 3. Fallback: comparar la parte antes de "_P#" con el nombre normalizado
      const i3 = headers.findIndex(h => {
        const hn = norm(h);
        const suffixPattern = new RegExp(`_P${periodo}$`);
        if (!suffixPattern.test(hn)) return false;
        const prefix = hn.replace(suffixPattern, '').replace(/\s+/g, '');
        return prefix === normAsig;
      });
      return i3;
    }

    // Promedio del periodo (desde hoja Calificaciones, columna "Promedio del periodo")
    function getNotaPeriodo(asignatura, periodo) {
      const entry = mapaCalif[norm(asignatura)] && mapaCalif[norm(asignatura)][String(periodo)];
      return entry ? entry.promedio : '—';
    }

    // Nota definitiva = max(promedio, recuperacion) — para desempeño y promedio general
    function getNotaDefinitiva(asignatura, periodo) {
      const entry = mapaCalif[norm(asignatura)] && mapaCalif[norm(asignatura)][String(periodo)];
      return entry ? entry.definitiva : '—';
    }

    // Recuperación de una asignatura en un periodo
    function getRecuperacion(asignatura, periodo) {
      const entry = mapaCalif[norm(asignatura)] && mapaCalif[norm(asignatura)][String(periodo)];
      return (entry && entry.recuperacion !== '—') ? entry.recuperacion : '—';
    }

    // Fallas de una asignatura en un periodo desde CONSOLIDADO_ASISTENCIA
    function getFallas(asignatura, periodo) {
      if (!filaAsist) return '—';
      const iCol = findColIdx(asistHeaders, asignatura, periodo);
      if (iCol < 0) return '—';
      const val = parseInt(filaAsist[iCol]);
      return isNaN(val) ? '0' : String(val);
    }

    // 8. Agrupar áreas
    const configFiltrada = configRows.filter(r => norm(r[iNivelConfig]) === norm(nivelEst));
    const areasMap = {};
    configFiltrada.forEach(row => {
      const area       = row[iAreaConfig]  || 'Sin área';
      const asignatura = row[iAsigConfig]  || '';
      const porcentaje = parseFloat(row[iPorcConfig])  || 100;
      const orden      = parseInt(row[iOrdenConfig])   || 999;
      if (!areasMap[area]) areasMap[area] = [];
      areasMap[area].push({ asignatura, porcentaje, orden });
    });
    Object.keys(areasMap).forEach(area => {
      areasMap[area].sort((a, b) => a.orden - b.orden);
    });

    // Encabezados de periodos
    let periodosHeader = '';
    for (let p = 1; p <= numPeriodo; p++) {
      periodosHeader += `<th style="border:1px solid #ccc; padding:6px 8px; text-align:center; width:52px; background:rgba(232,245,233,0.75); font-size:11px;">P${p}</th>`;
    }
    const totalCols = 1 + 1 + numPeriodo + 1 + 1;

    let tablaHTML = `
      <table style="width:100%; border-collapse:collapse; font-size:12px; margin-bottom:20px; line-height:1.25; background:transparent;">
        <thead>
          <tr>
            <th style="border:1px solid #ccc; padding:6px 10px; text-align:left; background:rgba(232,245,233,0.75); width:220px; font-size:11px;">ÁREA / ASIGNATURA</th>
            <th style="border:1px solid #ccc; padding:6px 6px; text-align:center; background:rgba(232,245,233,0.75); width:48px; font-size:11px;">Faltas</th>
            ${periodosHeader}
            <th style="border:1px solid #ccc; padding:6px 6px; text-align:center; background:rgba(232,245,233,0.75); width:52px; font-size:11px;">Recup.</th>
            <th style="border:1px solid #ccc; padding:6px 10px; text-align:center; background:rgba(232,245,233,0.75); width:80px; font-size:11px;">Desempeño</th>
          </tr>
        </thead>
        <tbody>`;

    Object.keys(areasMap).forEach(area => {
      const items       = areasMap[area];
      const esAreaUnica = items.length === 1 && items[0].porcentaje === 100;

      // Nota ponderada del área — usa nota definitiva (max promedio/recuperación)
      let notaPonderada = 0, totalPeso = 0;
      items.forEach(item => {
        const notaStr = getNotaDefinitiva(item.asignatura, numPeriodo);
        const nota    = parseFloat(notaStr) || 0;
        notaPonderada += nota * (item.porcentaje / 100);
        totalPeso     += item.porcentaje / 100;
      });
      const notaAreaNum = (totalPeso > 0 && notaPonderada > 0)
        ? Math.round((notaPonderada / totalPeso) * 10) / 10
        : null;
      const notaAreaStr = notaAreaNum !== null ? notaAreaNum.toFixed(1) : '—';

      // Recuperación del área (solo área única)
      const recupArea = esAreaUnica
        ? getRecuperacion(items[0].asignatura, numPeriodo)
        : '—';

      // Nota final del área = nota definitiva (ya calculada en getNotaDefinitiva)
      const notaFinalAreaStr = notaAreaStr;
      const desempenoArea    = calcularDesempeno(notaFinalAreaStr);

      // Docente del área
      let docenteArea = '';
      if (!esSedeUnitaria && filaEst) {
        const iColDoc = estHeaders.findIndex(h => norm(h) === norm(area));
        if (iColDoc !== -1) docenteArea = (filaEst[iColDoc] || '').trim();
      }

      // Fallas del área (suma de sus asignaturas)
      let fallasArea = 0, tieneFallasArea = false;
      items.forEach(item => {
        const f = getFallas(item.asignatura, numPeriodo);
        if (f !== '—') { fallasArea += parseInt(f) || 0; tieneFallasArea = true; }
      });
      const fallasAreaStr = tieneFallasArea ? String(fallasArea) : '—';

      // Fila del área
      tablaHTML += `
        <tr style="background:rgba(215,238,218,0.45);">
          <td style="border:1px solid #ccc; padding:2px 6px; font-weight:700; font-size:10.5px; line-height:1.2;">
            ${area}
            ${docenteArea ? `<span style="font-weight:500; font-size:11px; color:#2e7d32; display:block; margin-top:1px; line-height:1.2;">${docenteArea}</span>` : ''}
          </td>
          <td style="border:1px solid #ccc; padding:2px 5px; text-align:center; font-weight:600;">${fallasAreaStr}</td>`;

      for (let p = 1; p <= numPeriodo; p++) {
        const notaP = esAreaUnica ? getNotaPeriodo(items[0].asignatura, p) : (p === numPeriodo ? notaAreaStr : '—');
        tablaHTML += `<td style="border:1px solid #ccc; padding:2px 6px; text-align:center;
          font-weight:700; color:${notaP === '—' ? '#aaa' : 'inherit'};">${notaP}</td>`;
      }

      tablaHTML += `
          <td style="border:1px solid #ccc; padding:2px 5px; text-align:center; font-weight:600;">${recupArea}</td>
          <td style="border:1px solid #ccc; padding:2px 10px; text-align:center; font-weight:700; color:${desempenoArea.color};">
            ${notaFinalAreaStr !== '—' ? desempenoArea.texto : '—'}
          </td>
        </tr>`;

      // Indicador área única
      if (esAreaUnica) {
        const indicador = mapaIndicadores[norm(items[0].asignatura)] || '';
        if (indicador) tablaHTML += `
          <tr style="background:rgba(249,253,249,0.55);">
            <td colspan="${totalCols}" style="border:1px solid #ccc; padding:1.5px 12px 6px 16px;
                font-size:11px; font-style:italic; color:#172417; border-top:none;">
              <span style="color:#888; font-style:normal; font-weight:600; margin-right:4px;">▸</span>${indicador}
            </td>
          </tr>`;
      }

      // Sub-filas de asignaturas (área compuesta)
      if (!esAreaUnica) {
        items.forEach(item => {
          let docenteAsig = '';
          if (!esSedeUnitaria && filaEst) {
            const iColDoc = estHeaders.findIndex(h => norm(h) === norm(item.asignatura));
            if (iColDoc !== -1) docenteAsig = (filaEst[iColDoc] || '').trim();
          }

          const fallasAsig   = getFallas(item.asignatura, numPeriodo);
          const recupAsig    = getRecuperacion(item.asignatura, numPeriodo);
          const notaAsigFinal = getNotaDefinitiva(item.asignatura, numPeriodo);
          const desempenoAsig = calcularDesempeno(notaAsigFinal);

          tablaHTML += `
            <tr style="background:rgba(255,255,255,0.80);">
              <td style="border:1px solid #ccc; padding:1.5px 8px 2px 14px; font-size:11.5px; line-height:1.2;">
                <span style="font-weight:600;">${item.asignatura}</span>
                ${docenteAsig ? `<span style="font-weight:400; font-size:10.8px; color:#2e7d32; display:block; line-height:1.2;">${docenteAsig}</span>` : ''}
              </td>
              <td style="border:1px solid #ccc; padding:1.5px 5px; text-align:center;">${fallasAsig}</td>`;

          for (let p = 1; p <= numPeriodo; p++) {
            const notaP = getNotaPeriodo(item.asignatura, p);
            tablaHTML += `<td style="border:1px solid #ccc; padding:1.5px 5px; text-align:center;
              color:${notaP === '—' ? '#aaa' : 'inherit'};">${notaP}</td>`;
          }

          tablaHTML += `
              <td style="border:1px solid #ccc; padding:1.5px 5px; text-align:center;
                font-weight:${recupAsig !== '—' ? '600' : '400'};">${recupAsig}</td>
              <td style="border:1px solid #ccc; padding:1.5px 10px; text-align:center;
                font-weight:600; color:${desempenoAsig.color};">
                ${notaAsigFinal !== '—' ? desempenoAsig.texto : '—'}
              </td>
            </tr>`;

          const indicador = mapaIndicadores[norm(item.asignatura)] || '';
          if (indicador) tablaHTML += `
            <tr style="background:rgba(249,253,249,0.55);">
              <td colspan="${totalCols}" style="border:1px solid #ccc; padding:4px 12px 5px 22px;
                  font-size:11px; font-style:italic; color:#172417; border-top:none;">
                <span style="color:#888; font-style:normal; font-weight:600; margin-right:4px;">▸</span>${indicador}
              </td>
            </tr>`;
        });
      }
    });

    tablaHTML += `</tbody></table>`;

    const mensajeVacio = !filaCons ? `<div style="background:#fff3e0; border:1px solid #ff9800; color:#e65100; padding:12px; border-radius:8px; text-align:center; margin:12px 0;"><strong>⚠️ Sin calificaciones consolidadas aún</strong></div>` : '';

    let htmlBoletin = `
      <div style="max-width:100%; margin:0 auto; background:white; background-image:url('https://i.postimg.cc/4Np6dqtn/ESCUDO-IEAN-marca-de-agua.png'); background-repeat:no-repeat; background-position:center center; background-size:90%; padding:10px 12px; font-family:'DM Sans',Arial,sans-serif; font-size:12px; line-height:1.25; border:none; border-radius:0; box-shadow:none; position:relative; overflow:hidden;">
        <div style="position:relative; z-index:1;">
        <div style="display:flex; align-items:center; margin-bottom:6px;">
          <img src="https://i.postimg.cc/66rx6xzQ/ESCUDO-IEAN-baja-resolucion.png" style="height:60px; margin-right:10px;">
          <div style="flex:1;">
            <p style="margin:0; font-size:12px; font-weight:600; color:#1b5e20; line-height:1.2;">REPÚBLICA DE COLOMBIA - INSTITUCIÓN EDUCATIVA ANTONIO NARIÑO</p>
            <p style="margin:2px 0 0; font-size:11px; color:#444; line-height:1.2;">CÓDIGO DANE: 273152000142 CASABIANCA TOLIMA<br>RESOLUCIÓN DE APROBACIÓN N.° 5097 DEL 01 DE AGOSTO DE 2018</p>
            <h1 style="margin:4px 0 0; font-size:18px; color:#1b5e20; font-weight:700; line-height:1.1;">INFORME ACADÉMICO ${esFinalAnual ? '- FINAL ANUAL 2026' : `- PERIODO ${periodoSeleccionado} - 2026`}</h1>
          </div>
        </div>

        <table style="width:100%; border-collapse:collapse; margin-bottom:8px; font-size:12px; line-height:1.25;">
          <tr>
            <td style="border:1px solid #ddd; padding:6px 8px; font-size:12px;"><strong>Estudiante:</strong> ${nombreEst}</td>
            <td style="border:1px solid #ddd; padding:6px 8px; font-size:12px;"><strong>Documento:</strong> ${documento}</td>
            <td style="border:1px solid #ddd; padding:6px 8px; font-size:12px;"><strong>Grado:</strong> ${gradoEst}</td>
          </tr>
          <tr>
            <td style="border:1px solid #ddd; padding:6px 8px; font-size:12px;"><strong>Sede:</strong> ${sedeEst}</td>
            <td style="border:1px solid #ddd; padding:6px 8px; font-size:12px;" colspan="2"><strong>Director de Grado:</strong> ${directorGrado}</td>
          </tr>
        </table>

        ${mensajeVacio}
        ${tablaHTML}

        <!-- Promedio General y Ranking (justificado a la derecha) -->
        <div style="margin-top:6px; padding:6px 12px; background:rgba(232,245,233,0.95); border-radius:8px; text-align:right; font-size:12px;">
          <strong>Promedio General del Periodo:</strong> ${promedioGeneral} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <strong>Puesto:</strong> ${ranking}
        </div>

        <!-- Observaciones -->
        <div style="margin-top:8px; padding:7px 12px; background:rgba(248,249,250,0.95); border-radius:8px; border-left:5px solid #388e3c; min-height:80px;">
          <strong style="font-size:12px;">Observaciones generales:</strong><br><br>
          ${observacion || '<br>'}
        </div>

        <!-- Escala -->
        <div style="margin-top:4px; padding:5px 12px; background:rgba(241,248,233,0.95); border-radius:8px; font-size:10.5px; text-align:center;">
          <strong>ESCALA DE VALORACIÓN:</strong><br>
          BAJO (1,0 - 2,9) &nbsp;&nbsp; BÁSICO (3,0 - 3,9) &nbsp;&nbsp; ALTO (4,0 - 4,5) &nbsp;&nbsp; SUPERIOR (4,6 - 5,0)
        </div>

        <!-- Firma -->
        <div style="margin-top:60px; text-align:center;">
          <div style="border-top: 2px solid #333; width: 320px; margin: 0 auto 8px;"></div>
          <strong>${directorGrado}</strong><br>
          <span style="font-size:12px; color:#666;">Director de Grado</span>
        </div>

        <div style="text-align:center; margin-top:10px; color:#666; font-size:12px;">
          Sistema P.A.R.C.E • Institución Educativa Antonio Nariño • 2026
        </div>
        </div>
      </div>
    `;

    // Guardar en caché para impresión masiva
    if (typeof idx === 'number') cacheBoletin[idx] = htmlBoletin;
    const totalCargados = Object.keys(cacheBoletin).length;
    const barraPDF = document.getElementById('barra-pdf');
    if (barraPDF) {
      barraPDF.style.display = 'flex';
      document.getElementById('txt-cantidad-pdf').textContent = totalCargados;
    }

    // Si es modo silencioso (impresión masiva), no mostrar vista previa
    if (silencioso) return;

    document.getElementById('iframe-boletin').innerHTML = htmlBoletin;
    document.getElementById('vista-previa-boletin').style.display = 'block';
    document.getElementById('vista-previa-boletin').scrollIntoView({ behavior: 'smooth' });

  } catch (err) {
    console.error(err);
    alert("Error al generar el boletín.");
  }
}

// =============================================
//  FUNCIONES AUXILIARES (obligatorias)
// =============================================
function getNivelFromGrado(grado) {
  if (!grado) return 'Primaria';
  const g = grado.toLowerCase();
  if (['primero','segundo','tercero','cuarto','quinto'].some(n => g.includes(n))) return 'Primaria';
  if (['sexto','séptimo','octavo','noveno'].some(n => g.includes(n))) return 'Secundaria';
  if (['décimo','undécimo','10','11'].some(n => g.includes(n))) return 'Media';
  return 'Primaria';
}

function calcularDesempeno(nota) {
  const n = parseFloat(nota);
  if (isNaN(n)) return { texto: '—', color: '#666' };
  if (n >= 4.6) return { texto: 'SUPERIOR', color: '#2e7d32' };
  if (n >= 4.0) return { texto: 'ALTO',     color: '#66bb6a' };
  if (n >= 3.0) return { texto: 'BÁSICO',   color: '#f57c00' };
  return { texto: 'BAJO', color: '#e53935' };
}

// ════════════════════════════════════════════════════════
//  INICIALIZACIÓN
// ════════════════════════════════════════════════════════
async function init() {
  nombreAdmin = getUrlParameter('nombre') || 'Administrador';
  sedeAdmin   = getUrlParameter('sede')   || '';

  document.getElementById('sidebar-nombre').textContent = nombreAdmin;
  document.getElementById('sidebar-avatar').textContent =
    nombreAdmin.charAt(0).toUpperCase();

  // Navegación
  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.addEventListener('click', () => cambiarPagina(btn.dataset.page));
  });

  // Cargar filtros de boletines al iniciar
  cargarFiltrosBoletines();

  await cambiarPagina('dashboard');
}

async function cambiarPagina(page) {
  document.querySelectorAll('.nav-item').forEach(btn => btn.classList.remove('active'));
  const navBtn = document.querySelector(`[data-page="${page}"]`);
  if (navBtn) navBtn.classList.add('active');

  document.querySelectorAll('.page-content').forEach(p => p.classList.remove('active'));
  const pageEl = document.getElementById(`page-${page}`);
  if (pageEl) pageEl.classList.add('active');

  const titulos = {
    dashboard:   'Panel de Control',
    usuarios:    'Gestión de Usuarios',
    estudiantes: 'Estudiantes',
    periodos:    'Periodos Académicos',
    consolidado: 'Consolidado',
    sedes:       'Sedes y Grados',
    boletines:   'Generación de Boletines',
    avisos:      'Avisos'
  };
  document.getElementById('page-title').textContent = titulos[page] || 'Panel Administrador';

  // Cerrar sidebar en móvil al cambiar de página
  if (window.innerWidth <= 768) {
    document.getElementById('sidebar').classList.remove('open');
    document.getElementById('sidebar-overlay').classList.remove('open');
  }

  switch(page) {
    case 'dashboard':   await cargarDashboard();   break;
    case 'usuarios':    await cargarUsuarios();    break;
    case 'estudiantes': await cargarEstudiantes(); break;
    case 'periodos':    await cargarPeriodos();    break;
    case 'consolidado': await cargarConsolidado(); break;
    case 'sedes':       await cargarSedes();       break;
    case 'boletines':   await cargarFiltrosBoletines(); break;   // ← Importante
    case 'avisos':      await cargarAvisos();      break;
  }
}

// ════════════════════════════════════════════════════════
//  MÓDULO: SEDES Y GRADOS (Mejorado)
// ════════════════════════════════════════════════════════
async function cargarSedes() {
  try {
    // Cargar datos necesarios
    const [dataEst, dataListas] = await Promise.all([
      fetchSheet(SHEETS.estudiantes, true),
      fetchSheet('LISTAS_DESPLEGABLES', true)
    ]);

    if (!dataEst.values || dataEst.values.length < 2) {
      document.getElementById('lista-sedes').innerHTML = '<div class="empty-state"><p>No hay datos de sedes.</p></div>';
      document.getElementById('lista-grados').innerHTML = '<div class="empty-state"><p>No hay datos de grados.</p></div>';
      return;
    }

    const estHeaders = dataEst.values[0].map(h => h.trim());
    const estRows = dataEst.values.slice(1);

    const iSedeEst = estHeaders.findIndex(h => norm(h) === 'SEDE');
    const iGradoEst = estHeaders.findIndex(h => norm(h) === 'GRADO');
    const iEstadoEst = estHeaders.findIndex(h => norm(h) === 'ESTADO');

    // ==================== SEDES CON CÓDIGO DANE ====================
    let sedesHTML = '';

    if (dataListas.values && dataListas.values.length > 1) {
      const listasHeaders = dataListas.values[0].map(h => h.trim());
      const listasRows = dataListas.values.slice(1);

      const iSedeLista = listasHeaders.findIndex(h => norm(h) === 'SEDES');
      const iDane = listasHeaders.findIndex(h => norm(h) === 'DANE SEDE' || norm(h).includes('DANE'));

      // Obtener sedes únicas del colegio
      const sedesUnicas = [...new Set(
        estRows.map(r => (r[iSedeEst] || '').trim()).filter(Boolean)
      )].sort();

      sedesHTML = '<ul style="list-style:none; padding:0; margin:0;">';

      sedesUnicas.forEach(sede => {
        // Buscar código DANE
        let dane = '—';
        if (iSedeLista !== -1 && iDane !== -1) {
          const fila = listasRows.find(row => norm(row[iSedeLista] || '') === norm(sede));
          if (fila) dane = fila[iDane] || '—';
        }

        sedesHTML += `
          <li style="padding:12px 16px; border-bottom:1px solid #eee; display:flex; justify-content:space-between; align-items:center;">
            <strong>${sede}</strong>
            <span style="font-size:13px; color:#666; background:#f5f5f5; padding:4px 10px; border-radius:12px;">
              DANE: ${dane}
            </span>
          </li>`;
      });

      sedesHTML += '</ul>';
    } else {
      // Fallback si no hay lista desplegables
      const sedes = [...new Set(estRows.map(r => (r[iSedeEst] || '').trim()).filter(Boolean))].sort();
      sedesHTML = '<ul style="list-style:none; padding:16px 20px;">' +
        sedes.map(s => `<li style="padding:10px 0; border-bottom:1px solid #eee;"><strong>${s}</strong></li>`).join('') +
        '</ul>';
    }

    document.getElementById('lista-sedes').innerHTML = sedesHTML;

    // ==================== GRADOS CON CANTIDAD DE ESTUDIANTES ====================
    const conteoGrados = {};

    estRows.forEach(row => {
      const grado = (row[iGradoEst] || '').trim();
      const estado = (row[iEstadoEst] || '').trim().toLowerCase();

      if (grado && (estado === '' || estado === 'activo' || !['inactivo', 'retirado'].includes(estado))) {
        conteoGrados[grado] = (conteoGrados[grado] || 0) + 1;
      }
    });

    const gradosOrdenados = ordenarGrados(Object.keys(conteoGrados));

    let gradosHTML = '<ul style="list-style:none; padding:0; margin:0;">';

    gradosOrdenados.forEach(grado => {
      const cantidad = conteoGrados[grado];
      gradosHTML += `
        <li style="padding:12px 16px; border-bottom:1px solid #eee; display:flex; justify-content:space-between; align-items:center;">
          <strong>${grado}</strong>
          <span style="font-size:13px; color:#1b5e20; background:#e8f5e9; padding:4px 12px; border-radius:12px; font-weight:600;">
            ${cantidad} estudiantes
          </span>
        </li>`;
    });

    gradosHTML += '</ul>';

    document.getElementById('lista-grados').innerHTML = gradosHTML;

  } catch (err) {
    console.error('Error cargando sedes y grados:', err);
    document.getElementById('lista-sedes').innerHTML = '<div class="empty-state"><p>Error al cargar sedes.</p></div>';
    document.getElementById('lista-grados').innerHTML = '<div class="empty-state"><p>Error al cargar grados.</p></div>';
  }
}

// ════════════════════════════════════════════════════════
//  BOLETINES — CACHÉ HTML generado
// ════════════════════════════════════════════════════════
// Guarda el HTML de cada boletín ya generado: { idx: htmlString }
const cacheBoletin = {};

// Sobrescribir verPreviaBoletin para guardar en caché al generar
const _verPreviaBoletin_orig = verPreviaBoletin;
// (La caché se llena desde dentro de verPreviaBoletin — ver abajo)

// ════════════════════════════════════════════════════════
//  IMPRIMIR BOLETINES (uno o todos)
// ════════════════════════════════════════════════════════
async function imprimirBoletines(modo) {
  const periodo  = document.getElementById('filtro-periodo-boletin').value;
  const sede     = document.getElementById('filtro-sede-boletin').value;
  const grado    = document.getElementById('filtro-grado-boletin').value;

  if (!periodo) {
    showMsg('msg-boletines', 'Selecciona un periodo antes de imprimir.', 'error');
    return;
  }

  let htmlsAImprimir = [];

  if (modo === 'uno') {
    // Solo el boletín visible en la vista previa
    const contenido = document.getElementById('iframe-boletin').innerHTML;
    if (!contenido || contenido.trim() === '') {
      showMsg('msg-boletines', 'No hay ningún boletín en la vista previa.', 'error');
      return;
    }
    htmlsAImprimir = [contenido];

  } else {
    // Todos: generar los que faltan y recopilar todos
    showMsg('msg-boletines', `Preparando ${estudiantesBoletin.length} boletines...`, 'info');

    const btn = document.getElementById('btn-imprimir-todos');
    btn.disabled    = true;
    btn.textContent = 'Generando...';

    for (let i = 0; i < estudiantesBoletin.length; i++) {
      if (!cacheBoletin[i]) {
        await verPreviaBoletin(i, true); // true = modo silencioso
      }
      if (cacheBoletin[i]) htmlsAImprimir.push(cacheBoletin[i]);
    }

    btn.disabled    = false;
    btn.innerHTML   = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>
        <rect x="6" y="14" width="12" height="8"/></svg> Imprimir / PDF todos`;
  }

  if (!htmlsAImprimir.length) {
    showMsg('msg-boletines', 'No se pudieron generar los boletines.', 'error');
    return;
  }

  // Abrir ventana de impresión con todos los boletines, uno por página
  const ventana = window.open('', '_blank');
  ventana.document.write(`
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="UTF-8">
      <title>Boletines — I.E. Antonio Nariño</title>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        html, body { width: 100%; height: 100%; margin: 0; padding: 0; }
        body { font-family: 'DM Sans', Arial, sans-serif; background: white; }
        .pagina-boletin {
          page-break-after: always;
          page-break-inside: avoid;
          padding: 0;
          margin: 0;
        }
        .pagina-boletin:last-child { page-break-after: auto; }
        @page { size: letter portrait; margin: 1.5cm; }
        @media print {
          html, body { margin: 0; padding: 0; min-height: 100%; }
          .pagina-boletin { page-break-after: always; }
          .pagina-boletin:last-child { page-break-after: auto; }
          * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ccc; padding: 6px 10px; font-size: 12px; }
        img { max-height: 80px; }
      </style>
    </head>
    <body>
      ${htmlsAImprimir.map(h => `<div class="pagina-boletin">${h}</div>`).join('\n')}
      <script>
        window.onload = function() {
          setTimeout(function() { window.print(); }, 600);
        };
      <\/script>
    </body>
    </html>
  `);
  ventana.document.close();
}

// ════════════════════════════════════════════════════════
//  MÓDULO: AVISOS
// ════════════════════════════════════════════════════════

// Cache local de los avisos cargados (fila real en Sheets = índice + 2)
let avisosData = [];

// ── Cargar y renderizar lista ──────────────────────────
async function cargarAvisos() {
  const contenedor = document.getElementById('lista-avisos');
  contenedor.innerHTML = '<div class="module-loading"><div class="ring-loader"></div><p>Cargando avisos...</p></div>';

  try {
    const data = await fetchSheet(SHEETS.avisos, true);

    if (data.error || !data.values || data.values.length < 2) {
      avisosData = [];
      contenedor.innerHTML = `
        <div class="empty-state">
          <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
            <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/>
            <path d="M13.73 21a2 2 0 0 1-3.46 0"/>
          </svg>
          <p>Aún no hay avisos publicados.</p>
        </div>`;
      return;
    }

    avisosData = data.values;
    renderListaAvisos();

  } catch(err) {
    contenedor.innerHTML = '<div class="empty-state"><p>Error al cargar los avisos. Verifica tu conexión.</p></div>';
    console.error('Error cargarAvisos:', err);
  }
}

// ── Renderizar tabla de avisos ─────────────────────────
function renderListaAvisos() {
  const contenedor = document.getElementById('lista-avisos');
  if (!avisosData || avisosData.length < 2) {
    contenedor.innerHTML = `
      <div class="empty-state">
        <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/>
          <path d="M13.73 21a2 2 0 0 1-3.46 0"/>
        </svg>
        <p>Aún no hay avisos publicados.</p>
      </div>`;
    return;
  }

  const headers = avisosData[0].map(h => (h || '').trim());
  const rows    = avisosData.slice(1);

  const iTitulo = headers.findIndex(h => norm(h).includes('TITULO'));
  const iMensaje= headers.findIndex(h => norm(h).includes('MENSAJE'));
  const iDest   = headers.findIndex(h => norm(h).includes('DESTINATARIO') || norm(h).includes('DIRIGIDO'));
  const iFecha  = headers.findIndex(h => norm(h).includes('FECHA'));
  const iAutor  = headers.findIndex(h => norm(h).includes('AUTOR'));
  const iActivo = headers.findIndex(h => norm(h).includes('ACTIVO'));

  // Filtro activo del header
  const filtro = document.getElementById('filtro-estado-aviso')?.value || 'todos';

  const filasFiltradas = rows
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      const activo = norm(row[iActivo] || 'Si');
      const estaActivo = activo !== 'NO' && activo !== 'FALSE' && activo !== '0';
      if (filtro === 'activos')   return estaActivo;
      if (filtro === 'inactivos') return !estaActivo;
      return true;
    })
    .reverse(); // más recientes primero

  if (!filasFiltradas.length) {
    contenedor.innerHTML = '<div class="empty-state"><p>No hay avisos para el filtro seleccionado.</p></div>';
    return;
  }

  const colorDest = dest => {
    const d = norm(dest);
    if (d === 'TODOS')    return 'badge-info';
    if (d === 'DOCENTE' || d === 'DOCENTES') return 'badge-success';
    return 'badge-gray';
  };

  contenedor.innerHTML = `
    <div style="overflow-x:auto;">
      <table class="data-table">
        <thead>
          <tr>
            <th>Título</th>
            <th>Dirigido a</th>
            <th>Fecha</th>
            <th>Autor</th>
            <th>Estado</th>
            <th>Acciones</th>
          </tr>
        </thead>
        <tbody>
          ${filasFiltradas.map(({ row, idx }) => {
            const activo    = norm(row[iActivo] || 'Si');
            const estaActivo= activo !== 'NO' && activo !== 'FALSE' && activo !== '0';
            const titulo    = row[iTitulo]  || '—';
            const mensaje   = row[iMensaje] || '';
            const dest      = row[iDest]    || 'Todos';
            const fecha     = row[iFecha]   || '—';
            const autor     = row[iAutor]   || '—';
            const filaSheet = idx + 2; // fila real en Sheets (1 encabezado + 1 base)

            return `
            <tr class="${estaActivo ? '' : 'aviso-row-inactivo'}">
              <td>
                <div class="aviso-titulo-cell">
                  <strong>${titulo}</strong>
                  ${mensaje ? `<div class="aviso-preview">${mensaje.length > 80 ? mensaje.substring(0,80) + '…' : mensaje}</div>` : ''}
                </div>
              </td>
              <td><span class="badge ${colorDest(dest)}">${dest}</span></td>
              <td style="white-space:nowrap; color:var(--text-muted); font-size:12px;">${fecha}</td>
              <td style="font-size:12px;">${autor}</td>
              <td>
                <span class="badge ${estaActivo ? 'badge-success' : 'badge-danger'}">
                  ${estaActivo ? 'Activo' : 'Inactivo'}
                </span>
              </td>
              <td>
                <div style="display:flex; gap:6px; flex-wrap:wrap;">
                  <button class="btn btn-sm btn-outline"
                          onclick="toggleAviso(${filaSheet}, ${estaActivo})"
                          title="${estaActivo ? 'Desactivar aviso' : 'Activar aviso'}">
                    ${estaActivo ? 'Desactivar' : 'Activar'}
                  </button>
                  <button class="btn btn-sm"
                          onclick="eliminarAviso(${filaSheet}, '${titulo.replace(/'/g, "\\'")}')"
                          title="Eliminar aviso permanentemente"
                          style="background:#c62828; color:white; border:none;">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                      <polyline points="3 6 5 6 21 6"/>
                      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
                    </svg>
                  </button>
                </div>
              </td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
}

// ── Publicar aviso ─────────────────────────────────────
document.getElementById('btn-publicar-aviso').addEventListener('click', async () => {
  const titulo      = document.getElementById('av-titulo').value.trim();
  const mensaje     = document.getElementById('av-mensaje').value.trim();
  const destinatario= document.getElementById('av-destinatario').value;

  if (!titulo) {
    showMsg('msg-aviso', 'El título del aviso es obligatorio.', 'error');
    return;
  }
  if (!mensaje) {
    showMsg('msg-aviso', 'El mensaje no puede estar vacío.', 'error');
    return;
  }

  const hoy  = new Date();
  const fecha= `${String(hoy.getDate()).padStart(2,'0')}/${String(hoy.getMonth()+1).padStart(2,'0')}/${hoy.getFullYear()}`;

  const btn = document.getElementById('btn-publicar-aviso');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-sm"></span> Publicando...';

  try {
    const resp = await enviarAppsScript({
      tipo: 'publicarAviso',
      aviso: {
        titulo,
        mensaje,
        destinatario,
        fecha,
        autor: nombreAdmin || 'Administrador',
        activo: 'Si'
      }
    });

    if (resp.status === 'ok') {
      showMsg('msg-aviso', `✅ Aviso <strong>"${titulo}"</strong> publicado correctamente.`, 'success');
      // Limpiar formulario
      document.getElementById('av-titulo').value   = '';
      document.getElementById('av-mensaje').value  = '';
      document.getElementById('av-destinatario').value = 'Todos';
      // Refrescar lista
      dataCache[SHEETS.avisos] = null;
      await cargarAvisos();
    } else {
      showMsg('msg-aviso', `Error al publicar: ${resp.mensaje || 'Intenta de nuevo.'}`, 'error');
    }
  } catch(err) {
    showMsg('msg-aviso', 'Error de conexión. Verifica tu red e intenta de nuevo.', 'error');
    console.error(err);
  }

  btn.disabled = false;
  btn.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
    </svg> Publicar aviso`;
});

// ── Toggle activo/inactivo ─────────────────────────────
async function toggleAviso(filaSheet, estaActivo) {
  const nuevoEstado = estaActivo ? 'No' : 'Si';
  const texto       = estaActivo ? 'desactivar' : 'activar';

  if (!confirm(`¿Deseas ${texto} este aviso?`)) return;

  try {
    const resp = await enviarAppsScript({
      tipo:   'toggleAviso',
      fila:   filaSheet,
      activo: nuevoEstado
    });

    if (resp.status === 'ok') {
      dataCache[SHEETS.avisos] = null;
      await cargarAvisos();
    } else {
      alert('Error al cambiar el estado del aviso.');
    }
  } catch(err) {
    alert('Error de conexión.');
    console.error(err);
  }
}

// ── Eliminar aviso ─────────────────────────────────────
async function eliminarAviso(filaSheet, titulo) {
  if (!confirm(`¿Eliminar el aviso "${titulo}" permanentemente?\n\nEsta acción no se puede deshacer.`)) return;

  try {
    const resp = await enviarAppsScript({
      tipo: 'eliminarAviso',
      fila: filaSheet
    });

    if (resp.status === 'ok') {
      dataCache[SHEETS.avisos] = null;
      await cargarAvisos();
    } else {
      alert('Error al eliminar el aviso.');
    }
  } catch(err) {
    alert('Error de conexión.');
    console.error(err);
  }
}

// ── Filtro de estado ───────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  const filtroEl = document.getElementById('filtro-estado-aviso');
  if (filtroEl) filtroEl.addEventListener('change', renderListaAvisos);
});

document.addEventListener('DOMContentLoaded', init);
// ════════════════════════════════════════════════════════
//  ELIMINAR USUARIO
// ════════════════════════════════════════════════════════
function eliminarUsuario(filaIdx, nombreUsuario) {
  const nombre = nombreUsuario || "este usuario";

  if (!confirm(`¿Estás seguro de eliminar permanentemente al usuario "${nombre}"?`)) return;
  
  if (!confirm(`¡ÚLTIMA CONFIRMACIÓN!\n\nEsta acción es irreversible.\n\n¿Deseas eliminar a "${nombre}"?`)) return;

  enviarAppsScript({
    tipo: 'eliminarUsuario',
    fila: filaIdx + 1
  })
  .then(response => {
    if (response.status === "ok") {
      showMsg('msg-usuario', 
        `✅ El usuario <strong>${nombre}</strong> ha sido eliminado correctamente.`, 
        'success');
    } else {
      showMsg('msg-usuario', 'Error al eliminar el usuario.', 'error');
    }
    
    // Recargar tabla
    dataCache[SHEETS.usuarios] = null;
    cargarUsuarios();
  })
  .catch(err => {
    showMsg('msg-usuario', 'Error de conexión al eliminar el usuario.', 'error');
    console.error(err);
  });
}