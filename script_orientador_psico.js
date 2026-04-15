// ════════════════════════════════════════════════════════════════════
//  PARCE — Módulo Orientador Psicosocial  v3.0
//  Compatible con orientador_inicio.html v3 y style_orientador_psico.css v3
// ════════════════════════════════════════════════════════════════════
'use strict';

// ── CDN jsPDF + marked ────────────────────────────────────────────
(function(){
  if(!document.querySelector('script[src*="jspdf"]')){
    const s=document.createElement('script');
    s.src='https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
    document.head.appendChild(s);
  }
  if(!document.querySelector('script[src*="marked"]')){
    const m=document.createElement('script');
    m.src='https://cdnjs.cloudflare.com/ajax/libs/marked/9.1.6/marked.min.js';
    document.head.appendChild(m);
  }
})();

// ── Supabase ──────────────────────────────────────────────────────
const SB_URL = typeof SUPABASE_URL!=='undefined'?SUPABASE_URL:'';
const SB_KEY = typeof SUPABASE_ANON_KEY!=='undefined'?SUPABASE_ANON_KEY:'';

async function sbFetch(path, opts={}) {
  const res = await fetch(`${SB_URL}/rest/v1/${path}`, {
    ...opts,
    headers:{ 'apikey':SB_KEY,'Authorization':`Bearer ${SB_KEY}`,
      'Content-Type':'application/json','Prefer':opts.prefer||'return=representation',...opts.headers }
  });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  return res.status===204 ? null : res.json();
}

// ── Constantes Sheets (mismas que script_orientador.js si existe) ──
const _API_KEY        = "AIzaSyA1GsdvskVoIQDSOg6EI8aEQW5MnQsFeoE";
const _SPREADSHEET_ID = "1ytP_bXfbjRsA05KJYM8HAxQvTvnAeHgz09r4QhMy-Rs";
const _APPS_URL       = "https://script.google.com/macros/s/AKfycbygBydUcKQWczkcs4jL3BshWNY1npLtsqK9IbNA72nIgpRhmjHo3tKnCVry0AcwvGIf/exec";

function _fetchSheet(name){
  const url=`https://sheets.googleapis.com/v4/spreadsheets/${_SPREADSHEET_ID}/values/${encodeURIComponent(name)}?key=${_API_KEY}`;
  return fetch(url).then(r=>r.json());
}

// ── AES-256-GCM ───────────────────────────────────────────────────
const _SAL='PARCE-IEAN-ORI-2025';
async function _clave(uid){ const enc=new TextEncoder();
  const km=await crypto.subtle.importKey('raw',enc.encode(uid+_SAL),{name:'PBKDF2'},false,['deriveKey']);
  return crypto.subtle.deriveKey({name:'PBKDF2',salt:enc.encode(_SAL),iterations:100000,hash:'SHA-256'},km,{name:'AES-GCM',length:256},false,['encrypt','decrypt']); }
async function cifrar(txt,uid){ const k=await _clave(uid),iv=crypto.getRandomValues(new Uint8Array(12));
  const b=await crypto.subtle.encrypt({name:'AES-GCM',iv},k,new TextEncoder().encode(txt));
  const c=new Uint8Array(12+b.byteLength); c.set(iv); c.set(new Uint8Array(b),12);
  return btoa(String.fromCharCode(...c)); }
async function descifrar(b64,uid){ try{ const k=await _clave(uid),c=Uint8Array.from(atob(b64),x=>x.charCodeAt(0));
  return new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:c.slice(0,12)},k,c.slice(12)));
  }catch{ return '[Error al descifrar]'; } }

// ── Helpers ───────────────────────────────────────────────────────
function nrm(s){ return String(s||'').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function fFecha(iso){ if(!iso) return '—'; try{ return new Date(iso).toLocaleDateString('es-CO',{day:'2-digit',month:'short',year:'numeric'}); }catch{ return iso; } }
function setTxt(id,v){ const el=document.getElementById(id); if(el) el.textContent=v; }
function getUrlParameter(name) {
  name = name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]');
  const regex = new RegExp('[\\?&]' + name + '=([^&#]*)');
  const results = regex.exec(location.search);
  return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
}

const ORDEN_GRADOS=['Preescolar','Primero','Segundo','Tercero','Cuarto','Quinto','Sexto','Séptimo','Octavo','Noveno','Décimo','Undécimo'];
function sortGrados(arr){ return [...arr].sort((a,b)=>{ const ia=ORDEN_GRADOS.indexOf(a),ib=ORDEN_GRADOS.indexOf(b); if(ia===-1&&ib===-1) return a.localeCompare(b); if(ia===-1) return 1; if(ib===-1) return -1; return ia-ib; }); }

// ── Estado ────────────────────────────────────────────────────────
const psico={ oriId:'',oriNombre:'',sede:'',casos:[],citas:[],remisiones:[],estudiantes:[],casoExp:null,semanaOff:0 };

// ════════════════════════════════════════════════════════════════════
//  INIT
// ════════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', async ()=>{
  console.log('█ DOMContentLoaded - INICIO');
  
  psico.oriId      = getUrlParameter('correo')||getUrlParameter('usuario')||'orientador@iean.edu.co';
  psico.oriNombre  = getUrlParameter('nombre')||'Orientador';
  psico.sede       = getUrlParameter('sede')  ||'Sede Principal';
  console.log('█ Parámetros cargados');

  setTxt('header-nombre',  psico.oriNombre);
  setTxt('sidebar-nombre', psico.oriNombre);
  setTxt('header-sede-tag', psico.sede);
  console.log('█ Header actualizado');

  const ini=(psico.oriNombre[0]||'O').toUpperCase();
  ['header-avatar','sidebar-avatar'].forEach(id=>{ 
    const el=document.getElementById(id); 
    if(el&&!el.querySelector('img')) el.textContent=ini; 
  });
  console.log('█ Avatares listos');

  console.log('█ Llamando initNav()');
  initNav();
  console.log('█ initNav() completado');

  console.log('█ Inicializando modales...');
  try {
    initModales();
    console.log('█ Modales inicializados');
  } catch(e) {
    console.error('█ Error en initModales:', e);
  }

  console.log('█ Inicializando firma...');
  try {
    initFirma();
    console.log('█ Firma inicializada');
  } catch(e) {
    console.warn('█ Advertencia en firma:', e);
  }

  console.log('█ Inicializando monitor...');
  try {
    initMonitor();
    console.log('█ Monitor inicializado');
  } catch(e) {
    console.warn('█ Advertencia en monitor:', e);
  }

  // ── Cargar estudiantes en paralelo (necesario para los buscadores) ──
  console.log('█ Cargando estudiantes...');
  try {
    await esperarEstudiantes();
    console.log(`█ Estudiantes listos: ${psico.estudiantes.length}`);
    // Actualizar select de grados en el monitor ahora que tenemos datos
    const gs = document.getElementById('monitor-grado');
    if (gs && psico.estudiantes.length && gs.options.length <= 1) {
      const grados = sortGrados([...new Set(psico.estudiantes.map(e=>e.grado).filter(Boolean))]);
      grados.forEach(g=>{ const o=document.createElement('option'); o.value=g; o.textContent=g; gs.appendChild(o); });
    }
  } catch(e) {
    console.warn('█ Advertencia cargando estudiantes:', e);
  }

  console.log('█ DOMContentLoaded - FIN');
});

// ── Esperar todosEstudiantes de script_orientador.js ─────────────
async function esperarEstudiantes(){
  for(let i=0;i<30;i++){
    if(typeof todosEstudiantes!=='undefined'&&todosEstudiantes.length>0) break;
    await new Promise(r=>setTimeout(r,100));
  }
  if(typeof todosEstudiantes!=='undefined'&&todosEstudiantes.length>0){
    psico.estudiantes=todosEstudiantes.map(e=>({ documento:e.id||'',nombre:e.nombre||'',grado:e.grado||'',sede:e.sede||'' })).filter(e=>e.nombre);
  } else { await cargarEstudiantesDirecto(); }
}
async function cargarEstudiantesDirecto(){
  try{
    const data=await _fetchSheet('asignaturas_estudiantes');
    if(!data||data.error||!data.values||data.values.length<2) return;
    const h=data.values[0].map(x=>(x||'').trim());
    const col=ns=>h.findIndex(x=>ns.some(n=>nrm(x)===nrm(n)||nrm(x).includes(nrm(n))));
    const iId=col(['ID_ESTUDIANTE','DOCUMENTO','NO_DOCUMENTO']),iNm=col(['NOMBRES_APELLIDOS','NOMBRES Y APELLIDOS','NOMBRE']),iGr=col(['GRADO']),iSe=col(['SEDE']),iEs=col(['ESTADO']);
    psico.estudiantes=data.values.slice(1).filter(r=>{ if(iEs<0) return true; const e=nrm(r[iEs]||''); return e!=='INACTIVO'&&e!=='RETIRADO'; }).map(r=>({ documento:String(r[iId]||'').trim(),nombre:String(r[iNm]||'').trim(),grado:String(r[iGr]||'').trim(),sede:String(r[iSe]||'').trim() })).filter(e=>e.nombre);
  }catch(e){ console.error('[Psico] cargarEstudiantesDirecto:',e); }
}

// ── Supabase loaders ──────────────────────────────────────────────
async function cargarCasos(){ try{ psico.casos=await sbFetch('ori_casos?order=created_at.desc&limit=200',{method:'GET'})||[]; }catch(e){ psico.casos=[]; } }
async function cargarCitas(){ try{ psico.citas=await sbFetch('ori_citas?order=fecha_hora.asc&limit=200',{method:'GET'})||[]; }catch(e){ psico.citas=[]; } }
async function cargarRemisiones(){ try{ psico.remisiones=await sbFetch('ori_remisiones?order=created_at.desc&limit=200',{method:'GET'})||[]; }catch(e){ psico.remisiones=[]; } }

// ════════════════════════════════════════════════════════════════════
//  DASHBOARD
// ════════════════════════════════════════════════════════════════════
function actualizarDashboard(){
  const ab=psico.casos.filter(c=>c.estado!=='Cerrado');
  setTxt('sem-rojo-num',    ab.filter(c=>c.prioridad==='Rojo').length);
  setTxt('sem-naranja-num', ab.filter(c=>c.prioridad==='Naranja').length);
  setTxt('sem-verde-num',   ab.filter(c=>c.prioridad==='Verde').length);
  const hoy=new Date(), en7=new Date(hoy.getTime()+7*864e5);
  const proxC=psico.citas.filter(c=>{ const f=new Date(c.fecha_hora); return f>=hoy&&f<=en7&&c.estado==='Programada'; }).length;
  const sinR=psico.remisiones.filter(r=>r.estado_remision==='Sin Respuesta').length;
  setTxt('widget-casos-abiertos', ab.length);
  setTxt('widget-proximas-citas', proxC);
  setTxt('widget-remisiones-sin-resp', sinR);
  setTxt('widget-alertas-bienestar','—');

  // Badges nav
  const bd=document.getElementById('badge-dashboard'), bm=document.getElementById('badge-monitor'), br=document.getElementById('badge-remisiones');
  if(bd){ bd.textContent=ab.length; bd.style.display=ab.length?'':'none'; }
  if(br){ br.textContent=sinR; br.style.display=sinR?'':'none'; }

  renderCasos(psico.casos);
}

function renderCasos(casos, filtros={}){
  const cont=document.getElementById('casos-lista'), vacio=document.getElementById('casos-vacio');
  if(!cont) return;
  let lista=casos.filter(c=>c.estado!=='Cerrado');
  if(filtros.prioridad) lista=lista.filter(c=>c.prioridad===filtros.prioridad);
  if(filtros.estado)    lista=lista.filter(c=>c.estado===filtros.estado);
  if(!lista.length){ cont.innerHTML=''; if(vacio) vacio.style.display='flex'; return; }
  if(vacio) vacio.style.display='none';
  cont.innerHTML=lista.map(c=>`
    <div class="ori-caso-card" data-id="${c.id}">
      <div class="ori-dot dot-${(c.prioridad||'verde').toLowerCase()}"></div>
      <div class="ori-caso-info">
        <div class="ori-caso-nombre">${esc(c.estudiante_nombre)}</div>
        <div class="ori-caso-meta">${esc(c.grado||'')} · ${fFecha(c.fecha_apertura)}${c.motivo_apertura?' · '+esc(c.motivo_apertura.slice(0,50))+'…':''}</div>
      </div>
      <span class="ori-estado-badge estado-${eClass(c.estado)}">${esc(c.estado)}</span>
    </div>`).join('');
  cont.querySelectorAll('.ori-caso-card').forEach(card=>card.addEventListener('click',()=>abrirDetalle(card.dataset.id)));
}
const eClass=e=>({'Recibido':'recibido','En Valoración':'valoracion','Seguimiento Interno':'seguimiento','Remitido Externo':'remitido','Cerrado':'cerrado'}[e]||'recibido');

document.getElementById('filtro-casos-prioridad')?.addEventListener('change',e=>renderCasos(psico.casos,{prioridad:e.target.value,estado:document.getElementById('filtro-casos-estado')?.value}));
document.getElementById('filtro-casos-estado')?.addEventListener('change',e=>renderCasos(psico.casos,{estado:e.target.value,prioridad:document.getElementById('filtro-casos-prioridad')?.value}));

// ════════════════════════════════════════════════════════════════════
//  NAVEGACIÓN  (usa data-page / page-* / ori-nav-item / ori-page)
// ════════════════════════════════════════════════════════════════════
const PAGE_TITLES = {
  dashboard:'Dashboard de Alertas', expediente:'Expediente Integral',
  monitor:'Monitor de Riesgos', remisiones:'Remisiones Externas',
  agenda:'Agenda de Citas', manual:'Manual de Uso'
};

function initNav(){
  console.log('█ initNav() - INICIO');
  const navItems = document.querySelectorAll('.ori-nav-item');
  const pages    = document.querySelectorAll('.ori-page');
  const sb       = document.getElementById('ori-sidebar');
  const ov       = document.getElementById('ori-sidebar-overlay');

  console.log(`█ initNav: ${navItems.length} nav items, ${pages.length} pages`);
  
  navItems.forEach((item) => {
    item.addEventListener('click', () => {
      console.log(`█ NAV CLICK: ${item.dataset.page}`);
      const page = item.dataset.page;
      navItems.forEach(n=>n.classList.remove('ori-nav-active'));
      item.classList.add('ori-nav-active');
      pages.forEach(p=>p.classList.toggle('ori-page-active', p.id===`page-${page}`));
      setTxt('page-title', PAGE_TITLES[page]||'Orientación Psicosocial');
      console.log(`█ NAV CLICK completado para: ${page}`);
    });
  });

  document.getElementById('btn-menu-toggle')?.addEventListener('click',()=>{
    sb?.classList.toggle('sidebar-open'); ov?.classList.toggle('open');
  });
  ov?.addEventListener('click',()=>{ sb?.classList.remove('sidebar-open'); ov?.classList.remove('open'); });
  document.getElementById('btn-logout')?.addEventListener('click',()=>{
    if(confirm('¿Deseas cerrar sesión?')) window.location.href='index.html';
  });

  console.log('█ initNav() - FIN');
}

function navTo(page){
  document.querySelector(`.ori-nav-item[data-page="${page}"]`)?.click();
}

// ════════════════════════════════════════════════════════════════════
//  MANUAL
// ════════════════════════════════════════════════════════════════════
async function cargarManual(){
  const el=document.getElementById('manual-body');
  if(!el||el.dataset.cargado) return;
  try{
    const res=await fetch('manual_orientador_psicosocial.md');
    if(!res.ok) throw new Error('No encontrado');
    const txt=await res.text();
    el.innerHTML=typeof marked!=='undefined'?marked.parse(txt):`<pre style="white-space:pre-wrap;font-size:13px;">${esc(txt)}</pre>`;
    el.dataset.cargado='1';
  }catch(e){
    el.innerHTML='<p style="color:#aaa;padding:20px;">No se pudo cargar el manual. Asegúrate de que <code>manual_orientador_psicosocial.md</code> esté en la misma carpeta que el HTML.</p>';
  }
}

// ════════════════════════════════════════════════════════════════════
//  EXPEDIENTE
// ════════════════════════════════════════════════════════════════════
function resetExp(){
  const f=document.getElementById('exp-ficha-container');
  if(f) f.style.display='none';
  psico.casoExp=null;
  const i=document.getElementById('exp-buscar-input'); if(i) i.value='';
}

function initBuscadorExp(){
  const input=document.getElementById('exp-buscar-input'), sug=document.getElementById('exp-sugerencias');
  if(!input||!sug) return;
  input.addEventListener('input',()=>{
    // Re-intentar cargar si todavía vacío
    if(!psico.estudiantes.length&&typeof todosEstudiantes!=='undefined'&&todosEstudiantes.length){
      psico.estudiantes=todosEstudiantes.map(e=>({documento:e.id||'',nombre:e.nombre||'',grado:e.grado||'',sede:e.sede||''})).filter(e=>e.nombre);
    }
    const q=nrm(input.value); if(q.length<2){ sug.style.display='none'; return; }
    if(!psico.estudiantes.length){ sug.innerHTML='<div class="ori-dropdown-item" style="color:#aaa;">Sin estudiantes cargados</div>'; sug.style.display='block'; return; }
    const res=psico.estudiantes.filter(e=>nrm(e.nombre).includes(q)||nrm(e.documento).includes(q)).slice(0,8);
    if(!res.length){ sug.style.display='none'; return; }
    sug.innerHTML=res.map(e=>`<div class="ori-dropdown-item" data-doc="${esc(e.documento)}"><strong>${esc(e.nombre)}</strong> <span style="color:#aaa;font-size:11px;">${esc(e.grado)} · ${esc(e.documento)}</span></div>`).join('');
    sug.style.display='block';
    sug.querySelectorAll('.ori-dropdown-item').forEach(item=>item.addEventListener('click',()=>{ const est=psico.estudiantes.find(e=>e.documento===item.dataset.doc); if(est){ cargarExp(est); input.value=est.nombre; } sug.style.display='none'; }));
  });
  document.addEventListener('click',e=>{ if(!sug.contains(e.target)&&e.target!==input) sug.style.display='none'; });
}

async function cargarExp(est){
  const cont=document.getElementById('exp-ficha-container');
  if(!cont) return; cont.style.display='block';
  setTxt('exp-nombre',est.nombre);
  setTxt('exp-info-sub',`${est.grado} · Doc: ${est.documento}`);
  const av=document.getElementById('exp-avatar'); if(av) av.textContent=(est.nombre[0]||'?').toUpperCase();
  const ig=document.getElementById('exp-info-grid');
  if(ig) ig.innerHTML=[['Documento',est.documento],['Grado',est.grado],['Sede',est.sede]].map(([l,v])=>`<div class="ori-info-field"><span class="ori-info-label">${l}</span><span class="ori-info-value">${esc(v||'—')}</span></div>`).join('');
  const caso=psico.casos.find(c=>c.estudiante_id===est.documento&&c.estado!=='Cerrado');
  psico.casoExp=caso||null;
  if(caso){ await cargarTimeline(caso.id); initNotasPrivadas(caso.id); }
  else{ const tl=document.getElementById('exp-timeline'); if(tl) tl.innerHTML='<div class="ori-empty" style="padding:24px 20px;"><p>Sin caso activo. Ábrelo desde el Dashboard.</p></div>'; }
  await audit('VER_FICHA',{estudiante_id:est.documento,caso_id:caso?.id});
}

async function cargarTimeline(casoId){
  const cont=document.getElementById('exp-timeline'); if(!cont) return;
  cont.innerHTML='<div class="ori-loading" style="padding:16px 20px;"><div class="ori-spinner"></div><span>Cargando…</span></div>';
  try{
    const ints=await sbFetch(`ori_seguimientos?caso_id=eq.${casoId}&order=fecha_intervencion.desc`,{method:'GET'});
    if(!ints?.length){ cont.innerHTML='<div class="ori-empty" style="padding:24px 20px;"><p>Sin intervenciones registradas</p></div>'; return; }
    cont.innerHTML=ints.map(i=>`
      <div class="ori-tl-item">
        <div class="ori-tl-dot"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="4"/></svg></div>
        <div class="ori-tl-content">
          <div class="ori-tl-header"><span class="ori-tl-tipo">${esc(i.tipo_intervencion)}</span><span class="ori-tl-fecha">${fFecha(i.fecha_intervencion)}</span></div>
          <p class="ori-tl-desc">${esc(i.descripcion)}</p>
          ${i.notas_privadas?'<span class="ori-badge ori-badge-red" style="margin-top:4px;display:inline-block;">🔐 Nota privada</span>':''}
        </div>
      </div>`).join('');
  }catch(e){ cont.innerHTML='<div class="ori-empty" style="padding:24px 20px;"><p>Error cargando timeline</p></div>'; }
}

let _notasOk=false;
function initNotasPrivadas(casoId){
  if(_notasOk) return; _notasOk=true;
  document.getElementById('btn-toggle-notas')?.addEventListener('click',async()=>{
    const b=document.getElementById('notas-bloqueadas'),d=document.getElementById('notas-desbloqueadas'); if(!b||!d) return;
    if(d.style.display==='none'){
      b.style.display='none'; d.style.display='block';
      try{ const rows=await sbFetch(`ori_seguimientos?caso_id=eq.${casoId}&notas_privadas=not.is.null&order=created_at.desc&limit=1`,{method:'GET'});
        if(rows?.length&&rows[0].notas_privadas){ const t=document.getElementById('notas-privadas-texto'); if(t) t.value=await descifrar(rows[0].notas_privadas,psico.oriId); }
      }catch(e){}
      await audit('VER_NOTAS_PRIVADAS',{caso_id:casoId});
    }else{ b.style.display='flex'; d.style.display='none'; const t=document.getElementById('notas-privadas-texto'); if(t) t.value=''; }
  });
  document.getElementById('btn-guardar-notas')?.addEventListener('click',async()=>{
    const t=document.getElementById('notas-privadas-texto')?.value||'';
    if(!t.trim()){ toast('Escribe algo antes de guardar','error'); return; }
    if(!psico.casoExp){ toast('Sin caso activo','error'); return; }
    try{ const c=await cifrar(t,psico.oriId); await sbFetch('ori_seguimientos',{method:'POST',body:JSON.stringify({caso_id:psico.casoExp.id,tipo_intervencion:'Anotación Orientadora',descripcion:'[NOTA PRIVADA CIFRADA]',notas_privadas:c,orientador_id:psico.oriId,orientador_nombre:psico.oriNombre,estado_caso_al_registrar:psico.casoExp.estado})}); toast('✓ Notas cifradas y guardadas','success'); }
    catch(e){ toast('Error al guardar notas','error'); }
  });
}
document.getElementById('btn-nueva-intervencion')?.addEventListener('click',()=>{
  if(!psico.casoExp){ toast('Selecciona un estudiante con caso activo','error'); return; }
  document.getElementById('interv-caso-id').value=psico.casoExp.id;
  openM('modal-nueva-intervencion');
});

// ════════════════════════════════════════════════════════════════════
//  DETALLE CASO
// ════════════════════════════════════════════════════════════════════
const WF=['Recibido','En Valoración','Seguimiento Interno','Remitido Externo','Cerrado'];
async function abrirDetalle(casoId){
  const caso=psico.casos.find(c=>c.id===casoId); if(!caso) return;
  setTxt('detalle-caso-titulo',`Caso: ${caso.estudiante_nombre}`);
  const wf=document.getElementById('detalle-workflow');
  if(wf){ const idx=WF.indexOf(caso.estado);
    wf.innerHTML=WF.map((e,i)=>`<div class="ori-wf-step ${i<idx?'wf-done':''} ${i===idx?'wf-current':''}"><div class="ori-wf-dot">${i<idx?'✓':i+1}</div><div class="ori-wf-label">${e}</div></div>`).join(''); }
  const il=document.getElementById('detalle-intervenciones-lista');
  if(il){ il.innerHTML='<div class="ori-loading" style="padding:12px 0;"><div class="ori-spinner"></div><span>Cargando…</span></div>';
    try{ const ints=await sbFetch(`ori_seguimientos?caso_id=eq.${casoId}&order=fecha_intervencion.desc&limit=20`,{method:'GET'});
      il.innerHTML=(ints||[]).map(i=>`<div style="padding:10px 0;border-bottom:1px solid #f1f5f9;"><div style="display:flex;justify-content:space-between;"><strong style="font-size:13px;">${esc(i.tipo_intervencion)}</strong><span style="font-size:11px;color:#aaa;">${fFecha(i.fecha_intervencion)}</span></div><p style="font-size:13px;color:#555;margin-top:4px;">${esc(i.descripcion)}</p>${i.notas_privadas?'<span class="ori-badge ori-badge-red" style="margin-top:4px;display:inline-block;">🔐 Nota Privada</span>':''}</div>`).join('')||'<p style="color:#aaa;font-size:13px;padding:12px 0;">Sin intervenciones</p>';
    }catch{ il.innerHTML='<p style="color:#aaa;font-size:13px;padding:12px 0;">Error al cargar</p>'; } }
  const btn=document.getElementById('detalle-btn-intervencion');
  if(btn){ btn.onclick=()=>{ document.getElementById('interv-caso-id').value=casoId; closeM('modal-detalle-caso'); openM('modal-nueva-intervencion'); }; }
  openM('modal-detalle-caso');
  await audit('VER_FICHA',{caso_id:casoId,estudiante_id:caso.estudiante_id});
}

// ════════════════════════════════════════════════════════════════════
//  MONITOR
// ════════════════════════════════════════════════════════════════════
let _mData=[];
function initMonitor(){
  document.getElementById('btn-recargar-monitor')?.addEventListener('click',runMonitor);
  const gs=document.getElementById('monitor-grado');
  if(gs&&psico.estudiantes.length){
    const grados=sortGrados([...new Set(psico.estudiantes.map(e=>e.grado).filter(Boolean))]);
    grados.forEach(g=>{ const o=document.createElement('option'); o.value=g; o.textContent=g; gs.appendChild(o); });
  }
  ['monitor-grado','monitor-buscar'].forEach(id=>{ document.getElementById(id)?.addEventListener('input',filtrarM); document.getElementById(id)?.addEventListener('change',filtrarM); });
}

async function runMonitor(){
  const tbody=document.getElementById('monitor-tbody'),ld=document.getElementById('monitor-loading');
  if(!tbody) return; if(ld) ld.style.display='flex'; tbody.innerHTML='';
  try{
    const rc=await _fetchSheet('Calificaciones').catch(()=>null);
    const pa={};
    if(rc&&!rc.error&&rc.values?.length>1){
      const h=rc.values[0].map(x=>(x||'').trim());
      const iId=h.findIndex(x=>nrm(x)==='ID_ESTUDIANTE'), iPr=h.findIndex(x=>nrm(x).includes('PROMEDIO'));
      rc.values.slice(1).forEach(r=>{ const id=String(r[iId]||'').trim(),p=iPr>=0?parseFloat(String(r[iPr]||'').replace(',','.')):NaN; if(id&&!isNaN(p)&&p>0){ if(!pa[id]) pa[id]={s:0,c:0}; pa[id].s+=p; pa[id].c++; } });
    }
    _mData=psico.estudiantes.map(est=>{ const p=pa[est.documento]||{s:0,c:0},prom=p.c>0?p.s/p.c:null; return{...est,prom,rP:prom!==null&&prom<3.0}; }).filter(e=>e.rP);
    const bm=document.getElementById('badge-monitor');
    if(bm){ bm.textContent=_mData.length; bm.style.display=_mData.length?'':'none'; }
    filtrarM();
  }catch(e){ tbody.innerHTML=`<tr><td colspan="5" class="ori-tabla-vacia">Error: ${esc(e.message)}</td></tr>`; }
  finally{ if(ld) ld.style.display='none'; }
}

function filtrarM(){
  const tbody=document.getElementById('monitor-tbody'); if(!tbody) return;
  const gr=document.getElementById('monitor-grado')?.value||'';
  const bq=nrm(document.getElementById('monitor-buscar')?.value||'');
  let l=_mData;
  if(gr) l=l.filter(e=>e.grado===gr);
  if(bq) l=l.filter(e=>nrm(e.nombre).includes(bq));
  if(!l.length){ tbody.innerHTML='<tr><td colspan="5" class="ori-tabla-vacia">Sin estudiantes con promedio inferior a 3.0</td></tr>'; return; }
  l=[...l].sort((a,b)=>(a.prom??99)-(b.prom??99));
  tbody.innerHTML=l.map(e=>`<tr><td><strong>${esc(e.nombre)}</strong><br><span style="font-size:11px;color:#aaa;">${esc(e.documento)}</span></td><td>${esc(e.grado)}</td><td style="font-weight:700;color:#d97706;">${e.prom!==null?e.prom.toFixed(2):'—'}</td><td><span class="ori-risk-chip ori-risk-acad">📉 Bajo rendimiento</span></td><td><button class="ori-btn-ghost ori-btn-sm btn-m-abrir" data-doc="${esc(e.documento)}" data-nombre="${esc(e.nombre)}" data-grado="${esc(e.grado)}">Abrir Caso</button></td></tr>`).join('');
  tbody.querySelectorAll('.btn-m-abrir').forEach(btn=>btn.addEventListener('click',()=>{
    _caseEst={documento:btn.dataset.doc,nombre:btn.dataset.nombre,grado:btn.dataset.grado,sede:psico.sede};
    setTxt('caso-est-nombre-chip',btn.dataset.nombre);
    document.getElementById('caso-est-chip').style.display='flex';
    navTo('dashboard');
    setTimeout(()=>openM('modal-nuevo-caso'),200);
  }));
}

// ════════════════════════════════════════════════════════════════════
//  REMISIONES
// ════════════════════════════════════════════════════════════════════
function renderRemisiones(f='todos'){
  const c=document.getElementById('remisiones-lista'),v=document.getElementById('remisiones-vacio'); if(!c) return;
  let l=psico.remisiones; if(f!=='todos') l=l.filter(r=>r.estado_remision===f);
  if(!l.length){ c.innerHTML=''; if(v) v.style.display='flex'; return; } if(v) v.style.display='none';
  const ico={'EPS / IPS':'🏥','ICBF':'👶','Comisaría de Familia':'⚖️','Policía de Infancia':'👮','Fiscalía':'🏛️','Bienestar Social Municipal':'🏘️','Hospital / Urgencias':'🚑','Otro':'📋'};
  const bc={'Sin Respuesta':'ori-badge-red','En Proceso':'ori-badge-orange','Con Respuesta':'ori-badge-green','Cerrada':'ori-badge-gray'};
  c.innerHTML=l.map(r=>`<div class="ori-rem-card"><div class="ori-rem-icon">${ico[r.entidad_destino]||'📋'}</div><div class="ori-rem-info"><div class="ori-rem-title">${esc(r.entidad_destino)}</div><div class="ori-rem-meta">${fFecha(r.fecha_remision)}</div><p style="font-size:12px;color:#64748b;margin-top:4px;">${esc((r.descripcion_hechos||'').slice(0,100))}${(r.descripcion_hechos||'').length>100?'…':''}</p></div><div class="ori-rem-actions"><span class="ori-badge ${bc[r.estado_remision]||'ori-badge-gray'}">${esc(r.estado_remision)}</span><button class="ori-btn-ghost ori-btn-sm btn-pdf-r" data-id="${r.id}" style="margin-top:4px;">PDF</button></div></div>`).join('');
  c.querySelectorAll('.btn-pdf-r').forEach(btn=>btn.addEventListener('click',()=>{ const rem=psico.remisiones.find(r=>r.id===btn.dataset.id); if(rem) genPDF(rem); }));
}
document.querySelectorAll('[data-filter-rem]').forEach(pill=>pill.addEventListener('click',()=>{
  document.querySelectorAll('[data-filter-rem]').forEach(p=>p.classList.remove('ori-pill-active'));
  pill.classList.add('ori-pill-active'); renderRemisiones(pill.dataset.filterRem);
}));

// ════════════════════════════════════════════════════════════════════
//  AGENDA
// ════════════════════════════════════════════════════════════════════
function renderAgenda(){ renderSemana(); renderCitas(); }
function renderSemana(){
  const hB=new Date(); hB.setHours(0,0,0,0);
  const lu=new Date(hB); lu.setDate(lu.getDate()-((lu.getDay()+6)%7)+psico.semanaOff*7);
  const dias=Array.from({length:7},(_,i)=>{ const d=new Date(lu); d.setDate(d.getDate()+i); return d; });
  const dn=['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'],hs=hB.toDateString();
  const hdr=document.getElementById('agenda-semana-header');
  if(hdr) hdr.innerHTML=dias.map((d,i)=>`<div class="ori-dia-header ${d.toDateString()===hs?'dia-hoy':''}"><div class="ori-dia-nombre">${dn[i]}</div><div class="ori-dia-num">${d.getDate()}</div></div>`).join('');
  const sem=document.getElementById('agenda-citas-semana'); if(!sem) return;
  const sc=psico.citas.filter(c=>{ const f=new Date(c.fecha_hora); f.setHours(0,0,0,0); return f>=dias[0]&&f<=dias[6]; });
  sem.innerHTML=sc.length?sc.map(rCitaCard).join(''):'<div class="ori-empty" style="padding:24px 20px;"><p>Sin citas esta semana</p></div>';
}
function renderCitas(fe=''){
  const c=document.getElementById('citas-lista'),v=document.getElementById('citas-vacio'); if(!c) return;
  let l=[...psico.citas].sort((a,b)=>new Date(a.fecha_hora)-new Date(b.fecha_hora));
  if(fe) l=l.filter(x=>x.estado===fe);
  if(!l.length){ c.innerHTML=''; if(v) v.style.display='flex'; return; } if(v) v.style.display='none';
  c.innerHTML=l.map(rCitaCard).join('');
}
function rCitaCard(c){ const f=new Date(c.fecha_hora),h=f.toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'});
  const bc={'Programada':'ori-badge-blue','Realizada':'ori-badge-green','Cancelada':'ori-badge-gray','No Asistió':'ori-badge-orange'};
  return `<div class="ori-cita-card"><div class="ori-cita-hora">${h}</div><div class="ori-cita-info"><div class="ori-cita-nombre">${esc(c.estudiante_nombre)}</div><div class="ori-cita-meta">${esc(c.tipo_cita)} · ${esc(c.lugar||'Oficina de Orientación')}</div></div><span class="ori-badge ${bc[c.estado]||'ori-badge-gray'}">${esc(c.estado)}</span></div>`; }
document.getElementById('agenda-filtro-estado')?.addEventListener('change',e=>renderCitas(e.target.value));
document.getElementById('btn-semana-prev')?.addEventListener('click',()=>{ psico.semanaOff--; renderSemana(); });
document.getElementById('btn-semana-next')?.addEventListener('click',()=>{ psico.semanaOff++; renderSemana(); });
document.getElementById('btn-semana-hoy')?.addEventListener('click', ()=>{ psico.semanaOff=0; renderSemana(); });

// ════════════════════════════════════════════════════════════════════
//  MODALES
// ════════════════════════════════════════════════════════════════════
let _caseEst=null,_citaEst=null;
function initModales(){
  document.querySelectorAll('[data-close]').forEach(b=>b.addEventListener('click',()=>closeM(b.dataset.close)));
  document.querySelectorAll('.ori-modal-backdrop').forEach(bd=>bd.addEventListener('click',e=>{ if(e.target===bd) closeM(bd.id); }));

  document.getElementById('btn-nuevo-caso')?.addEventListener('click',()=>{ _caseEst=null; document.getElementById('caso-est-chip').style.display='none'; document.getElementById('caso-buscar-est').value=''; document.getElementById('caso-motivo').value=''; openM('modal-nuevo-caso'); });
  document.getElementById('btn-nueva-remision')?.addEventListener('click',()=>{ const sel=document.getElementById('rem-caso-select'); if(sel) sel.innerHTML='<option value="">Seleccionar caso…</option>'+psico.casos.filter(c=>c.estado!=='Cerrado').map(c=>`<option value="${c.id}">${esc(c.estudiante_nombre)} (${esc(c.estado)})</option>`).join(''); openM('modal-nueva-remision'); });
  document.getElementById('btn-nueva-cita')?.addEventListener('click',()=>{ _citaEst=null; document.getElementById('cita-est-chip').style.display='none'; document.getElementById('cita-buscar-est').value=''; const m=new Date(); m.setDate(m.getDate()+1); m.setHours(8,0,0,0); document.getElementById('cita-fecha-hora').value=m.toISOString().slice(0,16); openM('modal-nueva-cita'); });

  document.getElementById('rem-entidad')?.addEventListener('change',e=>{ const g=document.getElementById('rem-entidad-otro-group'); if(g) g.style.display=e.target.value==='Otro'?'block':'none'; });
  document.getElementById('interv-descripcion')?.addEventListener('input',e=>setTxt('interv-char-count',e.target.value.length));
  document.getElementById('interv-adjunto')?.addEventListener('change',e=>{ const f=e.target.files[0],p=document.getElementById('interv-adjunto-preview'); if(f&&p){ p.style.display='flex'; p.innerHTML=`📎 ${esc(f.name)} (${(f.size/1024).toFixed(0)} KB)`; } });

  initBuscadorModal('caso-buscar-est','caso-sugerencias',est=>{ _caseEst=est; setTxt('caso-est-nombre-chip',est.nombre); document.getElementById('caso-est-chip').style.display='flex'; document.getElementById('caso-buscar-est').value=''; });
  document.getElementById('caso-est-quitar')?.addEventListener('click',()=>{ _caseEst=null; document.getElementById('caso-est-chip').style.display='none'; });
  initBuscadorModal('cita-buscar-est','cita-sugerencias',est=>{ _citaEst=est; setTxt('cita-est-nombre-chip',est.nombre); document.getElementById('cita-est-chip').style.display='flex'; document.getElementById('cita-buscar-est').value=''; });
  document.getElementById('cita-est-quitar')?.addEventListener('click',()=>{ _citaEst=null; document.getElementById('cita-est-chip').style.display='none'; });

  document.getElementById('btn-guardar-caso')?.addEventListener('click',saveCaso);
  document.getElementById('btn-guardar-intervencion')?.addEventListener('click',saveInterv);
  document.getElementById('btn-guardar-remision')?.addEventListener('click',saveRem);
  document.getElementById('btn-exportar-pdf-remision')?.addEventListener('click',()=>genPDFModal());
  document.getElementById('btn-guardar-cita')?.addEventListener('click',saveCita);

  // Perfil — dos botones abren el mismo modal
  ['btn-mi-perfil','btn-mi-perfil-header'].forEach(id=>document.getElementById(id)?.addEventListener('click',abrirPerfil));
  document.getElementById('pfil-btn-guardar')?.addEventListener('click',guardarPerfil);

  initBuscadorExp();
}

function initBuscadorModal(inId,sgId,onS){
  const inp=document.getElementById(inId),sg=document.getElementById(sgId); if(!inp||!sg) return;
  inp.addEventListener('input',()=>{
    const q=nrm(inp.value); if(q.length<2){ sg.style.display='none'; return; }
    const r=psico.estudiantes.filter(e=>nrm(e.nombre).includes(q)||nrm(e.documento).includes(q)).slice(0,6);
    if(!r.length){ sg.style.display='none'; return; }
    sg.innerHTML=r.map(e=>`<div class="ori-dropdown-item" data-doc="${esc(e.documento)}"><strong>${esc(e.nombre)}</strong> <span style="color:#aaa;font-size:11px;">${esc(e.grado)}</span></div>`).join('');
    sg.style.display='block';
    sg.querySelectorAll('.ori-dropdown-item').forEach(item=>item.addEventListener('click',()=>{ const est=psico.estudiantes.find(e=>e.documento===item.dataset.doc); if(est) onS(est); sg.style.display='none'; }));
  });
  document.addEventListener('click',e=>{ if(!sg.contains(e.target)&&e.target!==inp) sg.style.display='none'; });
}

// ── Guardar Caso ──────────────────────────────────────────────────
async function saveCaso(){
  if(!_caseEst){ toast('Selecciona un estudiante','error'); return; }
  const pri=document.getElementById('caso-prioridad')?.value||'Verde',mot=document.getElementById('caso-motivo')?.value?.trim()||'';
  if(!mot){ toast('Escribe el motivo','error'); return; }
  const btn=document.getElementById('btn-guardar-caso'); btn.disabled=true; btn.textContent='Guardando…';
  try{
    const [nv]=await sbFetch('ori_casos',{method:'POST',body:JSON.stringify({estudiante_id:_caseEst.documento,estudiante_nombre:_caseEst.nombre,grado:_caseEst.grado,sede:_caseEst.sede||psico.sede,prioridad:pri,motivo_apertura:mot,orientador_id:psico.oriId,orientador_nombre:psico.oriNombre,estado:'Recibido'})});
    psico.casos.unshift(nv); actualizarDashboard(); closeM('modal-nuevo-caso');
    toast('✓ Caso abierto','success'); await audit('CREAR_CASO',{caso_id:nv.id,estudiante_id:_caseEst.documento}); _caseEst=null;
  }catch(e){ toast('Error al guardar','error'); console.error(e); }
  finally{ btn.disabled=false; btn.textContent='Abrir Caso'; }
}

// ── Guardar Intervención ──────────────────────────────────────────
async function saveInterv(){
  const cId=document.getElementById('interv-caso-id')?.value,tp=document.getElementById('interv-tipo')?.value,ds=document.getElementById('interv-descripcion')?.value?.trim(),ne=document.getElementById('interv-nuevo-estado')?.value;
  if(!cId){ toast('Sin caso seleccionado','error'); return; } if(!ds){ toast('Escribe la descripción','error'); return; }
  const btn=document.getElementById('btn-guardar-intervencion'); btn.disabled=true; btn.textContent='Guardando…';
  try{
    const caso=psico.casos.find(c=>c.id===cId);
    await sbFetch('ori_seguimientos',{method:'POST',body:JSON.stringify({caso_id:cId,tipo_intervencion:tp,descripcion:ds,orientador_id:psico.oriId,orientador_nombre:psico.oriNombre,estado_caso_al_registrar:caso?.estado||'Recibido'})});
    if(ne&&ne!==caso?.estado){ await sbFetch(`ori_casos?id=eq.${cId}`,{method:'PATCH',body:JSON.stringify({estado:ne})}); if(caso) caso.estado=ne; }
    const fi=document.getElementById('interv-adjunto'); if(fi?.files[0]) await subirAdj('seguimiento',cId,fi.files[0]);
    actualizarDashboard(); closeM('modal-nueva-intervencion'); toast('✓ Intervención guardada','success');
    await audit('CREAR_SEGUIMIENTO',{caso_id:cId});
    if(psico.casoExp?.id===cId) cargarTimeline(cId);
  }catch(e){ toast('Error al guardar','error'); console.error(e); }
  finally{ btn.disabled=false; btn.textContent='Guardar'; }
}

// ── Guardar Remisión ──────────────────────────────────────────────
async function saveRem(){
  const cId=document.getElementById('rem-caso-select')?.value,ent=document.getElementById('rem-entidad')?.value,mot=document.getElementById('rem-motivo-legal')?.value?.trim(),he=document.getElementById('rem-descripcion-hechos')?.value?.trim(),ac=document.getElementById('rem-acciones-previas')?.value?.trim();
  if(!cId){ toast('Selecciona un caso','error'); return; } if(!mot){ toast('Escribe el fundamento legal','error'); return; } if(!he){ toast('Escribe la descripción de hechos','error'); return; } if(!ac){ toast('Escribe las acciones previas','error'); return; }
  const btn=document.getElementById('btn-guardar-remision'); btn.disabled=true; btn.textContent='Guardando…';
  try{
    const canvas=document.getElementById('firma-canvas');
    const [nv]=await sbFetch('ori_remisiones',{method:'POST',body:JSON.stringify({caso_id:cId,entidad_destino:ent,entidad_nombre_especifico:document.getElementById('rem-entidad-otro')?.value||null,motivo_legal:mot,descripcion_hechos:he,acciones_previas_colegio:ac,orientador_id:psico.oriId,orientador_nombre:psico.oriNombre,firma_digital_base64:canvas?.toDataURL('image/png')||null})});
    psico.remisiones.unshift(nv); actualizarDashboard(); closeM('modal-nueva-remision');
    toast('✓ Remisión guardada','success'); await audit('CREAR_REMISION',{caso_id:cId}); renderRemisiones();
  }catch(e){ toast('Error al guardar remisión','error'); console.error(e); }
  finally{ btn.disabled=false; btn.textContent='Guardar Remisión'; }
}

// ── Guardar Cita ──────────────────────────────────────────────────
async function saveCita(){
  if(!_citaEst){ toast('Selecciona un estudiante','error'); return; }
  const tp=document.getElementById('cita-tipo')?.value,du=parseInt(document.getElementById('cita-duracion')?.value)||30,fh=document.getElementById('cita-fecha-hora')?.value,lu=document.getElementById('cita-lugar')?.value||'Oficina de Orientación',ds=document.getElementById('cita-descripcion')?.value?.trim();
  if(!fh){ toast('Selecciona fecha y hora','error'); return; }
  const btn=document.getElementById('btn-guardar-cita'); btn.disabled=true; btn.textContent='Guardando…';
  try{
    const caso=psico.casos.find(c=>c.estudiante_id===_citaEst.documento&&c.estado!=='Cerrado');
    const [nv]=await sbFetch('ori_citas',{method:'POST',body:JSON.stringify({caso_id:caso?.id||null,estudiante_id:_citaEst.documento,estudiante_nombre:_citaEst.nombre,tipo_cita:tp,fecha_hora:new Date(fh).toISOString(),duracion_minutos:du,lugar:lu,descripcion:ds||null,orientador_id:psico.oriId,estado:'Programada'})});
    psico.citas.push(nv); closeM('modal-nueva-cita'); toast('✓ Cita programada','success');
    renderAgenda(); actualizarDashboard(); await audit('CREAR_CITA',{estudiante_id:_citaEst.documento}); _citaEst=null;
  }catch(e){ toast('Error al guardar cita','error'); console.error(e); }
  finally{ btn.disabled=false; btn.textContent='Programar Cita'; }
}

// ════════════════════════════════════════════════════════════════════
//  PERFIL DEL ORIENTADOR
// ════════════════════════════════════════════════════════════════════
async function abrirPerfil(){
  openM('modal-perfil');
  document.getElementById('perfil-loading-ori').style.display='flex';
  document.getElementById('perfil-contenido-ori').style.display='none';
  document.getElementById('perfil-footer-ori').style.display='none';
  document.getElementById('pfil-feedback').style.display='none';

  try{
    const data=await _fetchSheet('Usuarios');
    if(data.error||!data.values||data.values.length<2){ showPfilFeedback('No se pudieron cargar los datos.','error'); document.getElementById('perfil-loading-ori').style.display='none'; return; }

    const headers=data.values[0].map(h=>(h||'').trim()), rows=data.values.slice(1);
    const iNom=headers.findIndex(h=>nrm(h).includes('NOMBRE')), iApel=headers.findIndex(h=>nrm(h).includes('APELLIDO'));
    const iDoc=headers.findIndex(h=>{ const n=nrm(h).replace(/[\s_\-]/g,''); return n==='DOCUMENTO'||n==='NODOCUMENTO'||n==='NUMERODOCUMENTO'||n==='CEDULA'||n==='NUMERODECEDULA'||n==='DOCIDENTIDAD'; });
    const iSede=headers.findIndex(h=>nrm(h)==='SEDE'), iRol=headers.findIndex(h=>nrm(h)==='ROL'||nrm(h)==='CARGO');
    const iUser=headers.findIndex(h=>nrm(h)==='USUARIO'), iPwd=headers.findIndex(h=>nrm(h)==='CONTRASENA');
    const iTel=headers.findIndex(h=>nrm(h)==='TELEFONO'||nrm(h)==='CELULAR');
    const iFoto=headers.findIndex(h=>nrm(h)==='FOTO_URL');

    let fila=null;
    for(const row of rows){
      const rn=nrm(String(row[iNom]||'')); let nc=rn;
      if(iApel>=0) nc=nrm(String(row[iNom]||'')+' '+String(row[iApel]||''));
      if(nc===nrm(psico.oriNombre)||rn===nrm(psico.oriNombre)||nrm(psico.oriNombre).includes(rn)||rn.includes(nrm(psico.oriNombre))){ fila=row; break; }
    }

    if(!fila){ showPfilFeedback('No se encontró tu registro.','error'); document.getElementById('perfil-loading-ori').style.display='none'; document.getElementById('perfil-contenido-ori').style.display='block'; document.getElementById('perfil-footer-ori').style.display='flex'; return; }

    const nombres=iNom>=0?String(fila[iNom]||'').trim():'—';
    const doc=iDoc>=0?String(fila[iDoc]||'').trim():'—';
    const sede=iSede>=0?String(fila[iSede]||'').trim():psico.sede;
    const rol=iRol>=0?String(fila[iRol]||'').trim():'Orientador';
    const usuario=iUser>=0?String(fila[iUser]||'').trim():'';
    const telefono=iTel>=0?String(fila[iTel]||'').trim():'';

    // Foto de perfil
    const fotoUrl=iFoto>=0?(typeof obtenerUrlFoto==='function'?obtenerUrlFoto(String(fila[iFoto]||'')):String(fila[iFoto]||'')):null;
    if(fotoUrl&&typeof crearWidgetFotoPerfil==='function'){
      crearWidgetFotoPerfil({ containerId:'perfil-foto-widget-ori', fotoUrl, iniciales:(nombres[0]||'O').toUpperCase(), tipoUsuario:'orientador', identificador:doc||psico.oriNombre,
        onSubida:async(url)=>{
          try{ await fetch(_APPS_URL,{method:'POST',body:JSON.stringify({tipo:'guardarFotoUrl',tipoUsuario:'orientador',nombre:psico.oriNombre,identificador:doc,fotoUrl:url})}); actualizarAvatarOri(url,nombres[0]); }catch(e){}
        }
      });
    }
    actualizarAvatarOri(fotoUrl, nombres[0]||'O');

    setTxt('pfil-nombre',nombres); setTxt('pfil-documento',doc); setTxt('pfil-sede',sede); setTxt('pfil-rol',rol);
    document.getElementById('pfil-usuario').value=usuario;
    document.getElementById('pfil-telefono').value=telefono;

    // Guardar referencia para guardar después
    window._pfilData={ headers, fila, doc, iUser, iPwd, iTel };

    document.getElementById('perfil-loading-ori').style.display='none';
    document.getElementById('perfil-contenido-ori').style.display='block';
    document.getElementById('perfil-footer-ori').style.display='flex';
  }catch(e){
    showPfilFeedback('Error de conexión.','error');
    document.getElementById('perfil-loading-ori').style.display='none';
    console.error(e);
  }
}

function actualizarAvatarOri(fotoUrl, inicial){
  ['header-avatar','sidebar-avatar'].forEach(id=>{
    const el=document.getElementById(id); if(!el) return;
    if(fotoUrl){ el.innerHTML=`<img src="${fotoUrl}" alt="Foto" style="width:100%;height:100%;object-fit:cover;border-radius:50%;">`; }
    else { el.textContent=(inicial||'O').toUpperCase(); }
  });
}

async function guardarPerfil(){
  const nuevoUsuario=document.getElementById('pfil-usuario')?.value.trim();
  const nuevoTelefono=document.getElementById('pfil-telefono')?.value.trim();
  const passAct=document.getElementById('pfil-passact')?.value;
  const passNew=document.getElementById('pfil-passnew')?.value;
  const passConf=document.getElementById('pfil-passconf')?.value;

  if(!nuevoUsuario){ showPfilFeedback('El campo de usuario no puede estar vacío.','error'); return; }
  let cambiarPass=false;
  if(passNew||passConf||passAct){
    if(!passAct){ showPfilFeedback('Ingresa tu contraseña actual para cambiarla.','error'); return; }
    if(passNew!==passConf){ showPfilFeedback('Las contraseñas no coinciden.','error'); return; }
    if(passNew.length<6){ showPfilFeedback('La contraseña debe tener al menos 6 caracteres.','error'); return; }
    cambiarPass=true;
  }

  const btn=document.getElementById('pfil-btn-guardar'); btn.disabled=true; btn.textContent='Guardando…';
  try{
    const payload={tipo:'actualizarPerfilDocente',nombre:psico.oriNombre,documento:document.getElementById('pfil-documento')?.textContent||'',nuevoUsuario,nuevoTelefono};
    if(cambiarPass){ payload.contrasenaActual=passAct; payload.nuevaContrasena=passNew; }
    const resp=await fetch(_APPS_URL,{method:'POST',body:JSON.stringify(payload)});
    const json=await resp.json();
    if(json.status==='ok'){
      showPfilFeedback('✓ Perfil actualizado correctamente.','ok');
      ['pfil-passact','pfil-passnew','pfil-passconf'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
    } else { showPfilFeedback(json.mensaje||'Error al guardar los cambios.','error'); }
  }catch(e){ showPfilFeedback('Error de conexión.','error'); }
  finally{ btn.disabled=false; btn.textContent='Guardar cambios'; }
}

function showPfilFeedback(msg,tipo){
  const el=document.getElementById('pfil-feedback'); if(!el) return;
  el.textContent=msg; el.style.display='block';
  el.style.background=tipo==='ok'?'#f0fdf4':'#fef2f2';
  el.style.color=tipo==='ok'?'#15803d':'#991b1b';
  el.style.border=tipo==='ok'?'1px solid #86efac':'1px solid #fca5a5';
  if(tipo==='ok') setTimeout(()=>el.style.display='none',5000);
}

// ════════════════════════════════════════════════════════════════════
//  FIRMA CANVAS
// ════════════════════════════════════════════════════════════════════
function initFirma(){
  const canvas=document.getElementById('firma-canvas'); if(!canvas) return;
  const ctx=canvas.getContext('2d'); let d=false;
  const gp=e=>{ const r=canvas.getBoundingClientRect(),sx=canvas.width/r.width,sy=canvas.height/r.height; return e.touches?{x:(e.touches[0].clientX-r.left)*sx,y:(e.touches[0].clientY-r.top)*sy}:{x:(e.clientX-r.left)*sx,y:(e.clientY-r.top)*sy}; };
  canvas.addEventListener('mousedown', e=>{ d=true; const p=gp(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); });
  canvas.addEventListener('mousemove', e=>{ if(!d) return; const p=gp(e); ctx.strokeStyle='#1d4ed8'; ctx.lineWidth=2; ctx.lineCap='round'; ctx.lineTo(p.x,p.y); ctx.stroke(); });
  canvas.addEventListener('mouseup',   ()=>d=false);
  canvas.addEventListener('mouseleave',()=>d=false);
  canvas.addEventListener('touchstart',e=>{ e.preventDefault(); d=true; const p=gp(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); },{passive:false});
  canvas.addEventListener('touchmove', e=>{ e.preventDefault(); if(!d) return; const p=gp(e); ctx.strokeStyle='#1d4ed8'; ctx.lineWidth=2; ctx.lineCap='round'; ctx.lineTo(p.x,p.y); ctx.stroke(); },{passive:false});
  canvas.addEventListener('touchend',  ()=>d=false);
  document.getElementById('btn-limpiar-firma')?.addEventListener('click',()=>ctx.clearRect(0,0,canvas.width,canvas.height));
}

// ════════════════════════════════════════════════════════════════════
//  PDF MEMBRETADO
// ════════════════════════════════════════════════════════════════════
async function genPDF(rem){
  const jsPDF=window.jspdf?.jsPDF||window.jsPDF;
  if(!jsPDF){ toast('Librería PDF cargando, intenta en 3 s…','error'); return; }
  const doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  const W=210,M=20,CW=W-2*M; let y=M;
  doc.setFillColor(15,23,42); doc.rect(0,0,W,28,'F');
  doc.setTextColor(255,255,255); doc.setFont('helvetica','bold'); doc.setFontSize(14);
  doc.text('I.E. ANTONIO NARIÑO',W/2,10,{align:'center'});
  doc.setFont('helvetica','normal'); doc.setFontSize(9);
  doc.text('Departamento de Orientación Escolar',W/2,16,{align:'center'});
  doc.text('La Dorada, Caldas',W/2,21,{align:'center'});
  y=36; doc.setTextColor(0,0,0);
  doc.setFont('helvetica','bold'); doc.setFontSize(13);
  doc.text('REMISIÓN A ENTIDAD EXTERNA',W/2,y,{align:'center'}); y+=6;
  doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(100,100,100);
  doc.text(`Fecha: ${new Date(rem.fecha_remision||Date.now()).toLocaleDateString('es-CO',{day:'2-digit',month:'long',year:'numeric'})}`,W/2,y,{align:'center'}); y+=10;
  doc.setDrawColor(59,130,246); doc.setLineWidth(0.5); doc.line(M,y,W-M,y); y+=6; doc.setTextColor(0,0,0);
  const sec=(t,tx)=>{ doc.setFontSize(10); doc.setFont('helvetica','bold'); doc.setTextColor(15,23,42); doc.text(t.toUpperCase(),M,y); y+=5; doc.setFont('helvetica','normal'); doc.setTextColor(40,40,40); doc.setFontSize(9); const ls=doc.splitTextToSize(tx||'—',CW); doc.text(ls,M,y); y+=ls.length*4.5+6; if(y>260){ doc.addPage(); y=M; } };
  sec('Entidad de Destino',rem.entidad_destino+(rem.entidad_nombre_especifico?` — ${rem.entidad_nombre_especifico}`:''));
  sec('Fundamento Legal',rem.motivo_legal); sec('Descripción de los Hechos',rem.descripcion_hechos); sec('Acciones Previas del Colegio',rem.acciones_previas_colegio);
  if(rem.firma_digital_base64?.startsWith('data:image')){ if(y>220){ doc.addPage(); y=M; } doc.setFontSize(9); doc.setTextColor(100,100,100); doc.text('Firma del Orientador:',M,y+2); doc.addImage(rem.firma_digital_base64,'PNG',M,y+5,60,20); y+=30; doc.line(M,y,M+60,y); y+=4; doc.text(rem.orientador_nombre||'Orientador Escolar',M,y); }
  const tp=doc.internal.getNumberOfPages();
  for(let p=1;p<=tp;p++){ doc.setPage(p); doc.setFontSize(8); doc.setTextColor(150,150,150); doc.text(`Pág. ${p}/${tp} · Generado por PARCE · Información Confidencial`,W/2,292,{align:'center'}); }
  doc.save(`Remision_${(rem.entidad_destino||'').replace(/\s/g,'_')}_${new Date(rem.fecha_remision||Date.now()).toISOString().slice(0,10)}.pdf`);
  await audit('EXPORTAR_PDF',{tipo:'remision'});
}
async function genPDFModal(){ await genPDF({ entidad_destino:document.getElementById('rem-entidad')?.value, entidad_nombre_especifico:document.getElementById('rem-entidad-otro')?.value, motivo_legal:document.getElementById('rem-motivo-legal')?.value, descripcion_hechos:document.getElementById('rem-descripcion-hechos')?.value, acciones_previas_colegio:document.getElementById('rem-acciones-previas')?.value, orientador_nombre:psico.oriNombre, firma_digital_base64:document.getElementById('firma-canvas')?.toDataURL('image/png'), fecha_remision:new Date().toISOString()}); }

// ── Adjuntos ──────────────────────────────────────────────────────
async function subirAdj(tipo,id,file){
  try{ const ext=file.name.split('.').pop(),path=`${tipo}/${id}/${Date.now()}.${ext}`; await fetch(`${SB_URL}/storage/v1/object/orientacion-adjuntos/${path}`,{method:'POST',headers:{'apikey':SB_KEY,'Authorization':`Bearer ${SB_KEY}`,'Content-Type':file.type},body:file}); await sbFetch('ori_adjuntos',{method:'POST',body:JSON.stringify({entidad_tipo:tipo,entidad_id:id,nombre_archivo:file.name,tipo_mime:file.type,storage_path:path,subido_por:psico.oriId})}); }catch(e){ console.error('[Adjunto]',e); }
}

// ── Auditoría ─────────────────────────────────────────────────────
async function audit(accion,det={}){
  try{ await sbFetch('ori_audit_log',{method:'POST',prefer:'return=minimal',body:JSON.stringify({accion,orientador_id:psico.oriId,orientador_nombre:psico.oriNombre,estudiante_id:det.estudiante_id||null,caso_id:det.caso_id||null,detalles:JSON.stringify(det),user_agent:navigator.userAgent.slice(0,200)})}); }catch(e){}
}

// ── UI helpers ────────────────────────────────────────────────────
function openM(id){ const m=document.getElementById(id); if(m) m.style.display='flex'; }
function closeM(id){ const m=document.getElementById(id); if(m) m.style.display='none'; }
let _tt=null;
function toast(msg,tipo='info'){ const t=document.getElementById('psico-toast'); if(!t) return; t.textContent=msg; t.className=`ori-toast toast-${tipo}`; t.style.display='block'; if(_tt) clearTimeout(_tt); _tt=setTimeout(()=>t.style.display='none',3500); }
