// ════════════════════════════════════════════════════════
//  P.A.R.C.E — Login v2
//  Autenticación dual:
//    · Docentes / Admin / Acudiente → hoja "Usuarios"
//    · Estudiantes → hoja "asignaturas_estudiantes"
//      Usuario   = columna ID_Estudiante  (col H)
//      Contraseña = Fecha_Nacimiento DD/MM/AAAA → sin barras → DDMMAAAA
// ════════════════════════════════════════════════════════

const API_KEY        = "AIzaSyA1GsdvskVoIQDSOg6EI8aEQW5MnQsFeoE";
const SPREADSHEET_ID = "1ytP_bXfbjRsA05KJYM8HAxQvTvnAeHgz09r4QhMy-Rs";

const URL_USUARIOS    = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/Usuarios?key=${API_KEY}`;
const URL_ESTUDIANTES = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/asignaturas_estudiantes?key=${API_KEY}`;
const URL_ESTUDIANTES_BASE = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/Estudiantes?key=${API_KEY}`;

// ── DOM ──
const loginForm    = document.getElementById('login-form');
const eyeToggle    = document.getElementById('eye-toggle');
const passwordInput= document.getElementById('password');
const btnLogin     = document.getElementById('btn-login');
const errorMsg     = document.getElementById('error-msg');
const errorText    = document.getElementById('error-text');
const forgotLink   = document.getElementById('forgot-link');

// ════════════════════════════════════════════════════════
//  1. OJO — alternar visibilidad contraseña
// ════════════════════════════════════════════════════════
eyeToggle.addEventListener('click', () => {
    const isText = passwordInput.type === 'text';
    passwordInput.type = isText ? 'password' : 'text';
    document.getElementById('eye-icon').innerHTML = isText
        ? '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>'
        : '<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/>';
});

// ════════════════════════════════════════════════════════
//  2. HELPERS
// ════════════════════════════════════════════════════════
function showError(msg) {
    errorText.textContent = msg;
    errorMsg.classList.add('visible');
}
function hideError() {
    errorMsg.classList.remove('visible');
}

/**
 * Normaliza la fecha de nacimiento como contraseña.
 * Acepta: "17/02/2010" → "17022010"
 *         "2010-02-17" → "17022010"
 *         "17022010"   → "17022010" (ya sin barras)
 */
function normalizarFecha(raw) {
    const s = String(raw || '').trim();
    // DD/MM/AAAA
    const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmy) return dmy[1].padStart(2,'0') + dmy[2].padStart(2,'0') + dmy[3];
    // AAAA-MM-DD
    const ymd = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (ymd) return ymd[3] + ymd[2] + ymd[1];
    // Ya sin barras (DDMMAAAA)
    return s.replace(/\D/g, '');
}

function norm(h) {
    return String(h||'').trim().toUpperCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}

// ════════════════════════════════════════════════════════
//  3. SUBMIT — flujo principal
// ════════════════════════════════════════════════════════
loginForm.addEventListener('submit', function(e) {
    e.preventDefault();
    hideError();

    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value.trim();

    if (!username || !password) {
        showError('Completa usuario y contraseña.');
        return;
    }

    btnLogin.classList.add('loading');
    btnLogin.disabled = true;

    // Consultar ambas hojas en paralelo
    Promise.all([
        fetch(URL_USUARIOS).then(r => r.json()).catch(() => ({ error: true })),
        fetch(URL_ESTUDIANTES).then(r => r.json()).catch(() => ({ error: true }))
    ]).then(([dataUsuarios, dataEstudiantes]) => {
        btnLogin.classList.remove('loading');
        btnLogin.disabled = false;

        // ── A. Buscar en hoja Usuarios (Docentes / Admin / Acudiente) ──
        if (!dataUsuarios.error && dataUsuarios.values && dataUsuarios.values.length > 1) {
            const [headers, ...rows] = dataUsuarios.values;
            const cleanH = headers.map(h => h.trim());
            for (const row of rows) {
                const u = {};
                row.forEach((v, i) => { u[cleanH[i]] = (v || '').trim(); });
                if (u['Usuario'] === username && u['Contraseña'] === password) {
                    const activo = norm(u['Activo'] || u['Estado'] || 'SI');
                    if (activo === 'NO' || activo === 'INACTIVO' || activo === 'FALSE') {
                        showError('Tu cuenta está desactivada. Comunícate con el administrador.');
                        return;
                    }
                    mostrarBienvenida(u['Nombre'] || username, u['Rol'] || '', () => {
                        procesarRedireccion(u);
                    });
                    return;
                }
            }
        }

        // ── B. Buscar en hoja asignaturas_estudiantes ──
        if (!dataEstudiantes.error && dataEstudiantes.values && dataEstudiantes.values.length > 1) {
            const [headers, ...rows] = dataEstudiantes.values;
            const cleanH = headers.map(h => h.trim());

            // Columnas en asignaturas_estudiantes
            const iId     = cleanH.findIndex(h => norm(h) === 'ID_ESTUDIANTE');
            const iPwd    = cleanH.findIndex(h => norm(h) === 'CONTRASENA');  // columna Contraseña (norm convierte Ñ→N)
            const iNombre = cleanH.findIndex(h => norm(h).includes('NOMBRE'));
            const iGrado  = cleanH.findIndex(h => norm(h) === 'GRADO');
            const iSede   = cleanH.findIndex(h => norm(h) === 'SEDE');
            const iActivo = cleanH.findIndex(h => norm(h) === 'ACTIVO' || norm(h) === 'ESTADO');

            if (iId < 0) {
                showError('Configuración incompleta. Contacta al administrador.');
                return;
            }

            for (const row of rows) {
                const id          = String(row[iId]  || '').trim();
                const claveHoja   = iPwd >= 0 ? String(row[iPwd] || '').trim() : '';
                // La columna Contraseña puede traer fecha con formato DD/MM/AAAA (desde BUSCARX)
                // o una clave personalizada en texto plano
                const claveNorm   = normalizarFecha(claveHoja); // si es fecha la normaliza, si no, la deja igual

                if (id !== username) continue;

                // Comparar: primero intenta con la clave normalizada (cubre fecha DD/MM/AAAA)
                // luego con la clave tal cual (cubre contraseñas personalizadas en texto plano)
                const coincide = (claveNorm === password) || (claveHoja === password);
                if (!coincide) continue;

                const activo = iActivo >= 0 ? norm(String(row[iActivo] || 'SI')) : 'SI';
                if (activo === 'NO' || activo === 'INACTIVO' || activo === 'FALSE') {
                    showError('Tu cuenta está inactiva. Consulta a tu rector/a.');
                    return;
                }

                const nombre = iNombre >= 0 ? String(row[iNombre] || '').trim() : username;
                const grado  = iGrado  >= 0 ? String(row[iGrado]  || '').trim() : '';
                const sede   = iSede   >= 0 ? String(row[iSede]   || '').trim() : '';

                mostrarBienvenida(nombre, 'Estudiante', () => {
                    const params = `?nombre=${encodeURIComponent(nombre)}&sede=${encodeURIComponent(sede)}&grado=${encodeURIComponent(grado)}`;
                    window.location.href = 'estudiante_inicio.html' + params;
                });
                return;
            }
        }

        // ── No encontrado ──
        showError('Usuario o contraseña incorrectos. Verifica tus datos.');

    }).catch(err => {
        btnLogin.classList.remove('loading');
        btnLogin.disabled = false;
        showError('Error de conexión. Verifica tu internet.');
        console.error(err);
    });
});

// ════════════════════════════════════════════════════════
//  4. REDIRECCIÓN por rol (Docente / Admin / Acudiente)
// ════════════════════════════════════════════════════════
function procesarRedireccion(user) {
    const nombre = user['Nombre'] || '';
    const sede   = user['Sede']   || '';
    const rol    = user['Rol']    || '';
    const base   = `?nombre=${encodeURIComponent(nombre)}&sede=${encodeURIComponent(sede)}`;

    const rutas = {
        'Administrador': 'admin_inicio.html',
        'Docente':       'docente_inicio.html',
        'Acudiente':     'acudiente_inicio.html',
    };

    if (rutas[rol]) {
        window.location.href = rutas[rol] + base;
    } else if (rol === 'Estudiante') {
        const grado = user['Grado'] || '';
        window.location.href = `estudiante_inicio.html${base}&grado=${encodeURIComponent(grado)}`;
    } else {
        showError('Rol no reconocido: ' + rol);
    }
}

// ════════════════════════════════════════════════════════
//  5. BIENVENIDA — overlay animado antes de redirigir
// ════════════════════════════════════════════════════════
function mostrarBienvenida(nombre, rol, callback) {
    // Crear overlay
    const overlay = document.createElement('div');
    overlay.id = 'welcome-overlay';
    overlay.innerHTML = `
      <div class="welcome-box">
        <div class="welcome-avatar">${(nombre[0] || '?').toUpperCase()}</div>
        <div class="welcome-greeting">¡Bienvenido/a!</div>
        <div class="welcome-name">${primerNombre(nombre)}</div>
        <div class="welcome-rol">${rolLabel(rol)}</div>
        <div class="welcome-bar"><div class="welcome-progress"></div></div>
      </div>`;
    document.body.appendChild(overlay);

    // Animar barra
    requestAnimationFrame(() => {
        overlay.querySelector('.welcome-progress').style.width = '100%';
    });

    setTimeout(callback, 1800);
}

function primerNombre(nombre) {
    return (nombre || '').split(' ')[0] || nombre;
}

function rolLabel(rol) {
    const labels = {
        'Administrador': '👨‍💼 Administrador',
        'Docente':       '👩‍🏫 Docente',
        'Estudiante':    '🎒 Estudiante',
        'Acudiente':     '👨‍👧 Acudiente',
    };
    return labels[rol] || rol;
}

// ════════════════════════════════════════════════════════
//  6. PANTALLA "¿OLVIDASTE TU CONTRASEÑA?" → INFO MODAL
// ════════════════════════════════════════════════════════
if (forgotLink) {
    forgotLink.addEventListener('click', (e) => {
        e.preventDefault();
        mostrarModalAyuda();
    });
}

function mostrarModalAyuda() {
    const modal = document.createElement('div');
    modal.id = 'modal-ayuda';
    modal.innerHTML = `
      <div class="modal-box">
        <button class="modal-close" id="modal-close-btn">✕</button>
        <div class="modal-icon">🔑</div>
        <h3 class="modal-title">¿Olvidaste tu contraseña?</h3>
        <div class="modal-body">
          <div class="modal-tip">
            <div class="tip-icon">🎒</div>
            <div>
              <strong>Si eres Estudiante:</strong><br>
              Tu usuario es tu <b>número de documento</b> de identidad.<br>
              Tu contraseña es tu <b>fecha de nacimiento</b> sin barras.<br>
              <span class="tip-example">Ejemplo: naciste el 17/02/2010 → contraseña: <code>17022010</code></span>
            </div>
          </div>
          <div class="modal-tip">
            <div class="tip-icon">👩‍🏫</div>
            <div>
              <strong>Si eres Docente o Administrativo:</strong><br>
              Comunícate con el administrador del sistema para que restablezca tu contraseña.
            </div>
          </div>
        </div>
        <button class="modal-btn" id="modal-close-btn2">Entendido</button>
      </div>`;
    document.body.appendChild(modal);

    const cerrar = () => modal.remove();
    document.getElementById('modal-close-btn').addEventListener('click', cerrar);
    document.getElementById('modal-close-btn2').addEventListener('click', cerrar);
    modal.addEventListener('click', (e) => { if (e.target === modal) cerrar(); });
}

// ════════════════════════════════════════════════════════
//  7. ESTILOS DINÁMICOS (overlay y modal)
// ════════════════════════════════════════════════════════
const dynamicStyles = document.createElement('style');
dynamicStyles.textContent = `

/* ── BIENVENIDA ── */
#welcome-overlay {
  position: fixed; inset: 0; z-index: 9999;
  background: linear-gradient(145deg, #2e7d32, #1b5e20);
  display: flex; align-items: center; justify-content: center;
  animation: fadeIn .3s ease;
}
@keyframes fadeIn { from { opacity:0 } to { opacity:1 } }

.welcome-box {
  text-align: center; color: #fff; padding: 40px 32px;
  display: flex; flex-direction: column; align-items: center; gap: 8px;
}
.welcome-avatar {
  width: 72px; height: 72px; border-radius: 50%;
  background: rgba(255,255,255,.2);
  font-size: 32px; font-weight: 800;
  display: flex; align-items: center; justify-content: center;
  border: 3px solid rgba(255,255,255,.4);
  margin-bottom: 8px;
}
.welcome-greeting { font-size: 15px; opacity: .8; }
.welcome-name {
  font-family: 'Playfair Display', serif;
  font-size: 30px; font-weight: 600; line-height: 1.1;
}
.welcome-rol {
  font-size: 14px; opacity: .75;
  background: rgba(255,255,255,.15);
  padding: 4px 16px; border-radius: 99px; margin-top: 4px;
}
.welcome-bar {
  width: 200px; height: 4px;
  background: rgba(255,255,255,.2); border-radius: 99px;
  margin-top: 24px; overflow: hidden;
}
.welcome-progress {
  height: 100%; width: 0; background: #fff; border-radius: 99px;
  transition: width 1.6s cubic-bezier(.4,0,.2,1);
}

/* ── MODAL AYUDA ── */
#modal-ayuda {
  position: fixed; inset: 0; z-index: 9998;
  background: rgba(0,0,0,.45);
  display: flex; align-items: center; justify-content: center;
  padding: 20px;
  animation: fadeIn .2s ease;
}
.modal-box {
  background: #fff; border-radius: 20px;
  padding: 32px 28px; max-width: 380px; width: 100%;
  position: relative;
  box-shadow: 0 20px 60px rgba(0,0,0,.2);
}
.modal-close {
  position: absolute; top: 14px; right: 16px;
  background: none; border: none; font-size: 16px;
  cursor: pointer; color: #999; padding: 4px 8px;
}
.modal-icon { font-size: 36px; text-align: center; margin-bottom: 8px; }
.modal-title {
  font-family: 'Playfair Display', serif;
  font-size: 22px; text-align: center; margin-bottom: 20px; color: #1a2e1a;
}
.modal-body { display: flex; flex-direction: column; gap: 16px; margin-bottom: 24px; }
.modal-tip {
  display: flex; gap: 12px; align-items: flex-start;
  background: #f0f4f0; border-radius: 12px; padding: 14px;
  font-size: 14px; line-height: 1.55; color: #3d5c3d;
}
.tip-icon { font-size: 22px; flex-shrink: 0; margin-top: 2px; }
.tip-example {
  display: block; margin-top: 6px; font-size: 13px;
  color: #2e7d32;
}
.tip-example code {
  background: #e8f5e9; padding: 2px 8px; border-radius: 6px;
  font-family: monospace; font-weight: 700;
}
.modal-btn {
  width: 100%; padding: 14px;
  background: #2e7d32; color: #fff;
  border: none; border-radius: 12px;
  font-size: 15px; font-weight: 600; cursor: pointer;
  transition: background .2s;
}
.modal-btn:hover { background: #1b5e20; }

/* ── CAMBIAR CONTRASEÑA (en portal estudiante) ── */
.cambiar-pwd-card {
  background: #fff; border-radius: 18px;
  padding: 20px; margin: 12px 0;
  box-shadow: 0 2px 8px rgba(0,0,0,.06);
}
.cambiar-pwd-title {
  font-size: 16px; font-weight: 800; margin-bottom: 14px; color: #1a1d2e;
}
.pwd-group { margin-bottom: 14px; }
.pwd-group label { display: block; font-size: 12px; font-weight: 700; color: #6b7280; margin-bottom: 6px; text-transform: uppercase; }
.pwd-input-wrap { position: relative; }
.pwd-input {
  width: 100%; padding: 13px 44px 13px 14px;
  border: 2px solid #e5e7eb; border-radius: 12px;
  font-size: 16px; font-family: 'Nunito', sans-serif;
  outline: none; transition: border-color .2s;
}
.pwd-input:focus { border-color: #4c6ef5; }
.pwd-eye {
  position: absolute; right: 12px; top: 50%; transform: translateY(-50%);
  background: none; border: none; cursor: pointer; color: #9ca3af;
}
.pwd-hint { font-size: 12px; color: #9ca3af; margin-top: 4px; }
.pwd-error { font-size: 13px; color: #c92a2a; margin-top: 8px; display: none; }
.pwd-error.show { display: block; }
.btn-cambiar-pwd {
  width: 100%; padding: 13px;
  background: #4c6ef5; color: #fff;
  border: none; border-radius: 12px;
  font-family: 'Nunito', sans-serif;
  font-size: 15px; font-weight: 800; cursor: pointer;
  transition: background .2s;
}
.btn-cambiar-pwd:active { background: #3b5bdb; }
.pwd-success {
  background: #d3f9d8; border: 2px solid #40c057;
  border-radius: 12px; padding: 12px; text-align: center;
  font-size: 14px; font-weight: 700; color: #2f9e44;
  display: none;
}
.pwd-success.show { display: block; }
`;
document.head.appendChild(dynamicStyles);
