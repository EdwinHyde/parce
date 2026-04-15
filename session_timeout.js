// ════════════════════════════════════════════════════════
//  P.A.R.C.E — Session Timeout Manager
//  Cierre automático de sesión tras 30 min de inactividad
// ════════════════════════════════════════════════════════

(function() {
  'use strict';

  // ── CONFIGURACIÓN ──
  const INACTIVITY_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutos
  const WARNING_BEFORE_MS     = 1 * 60 * 1000;  // Aviso 1 min antes
  const STORAGE_KEY           = 'parce_session_timer';

  let inactivityTimer  = null;
  let warningTimer     = null;
  let warningModal     = null;

  // ════════════════════════════════════════════════════════
  //  FUNCIONES PRINCIPALES
  // ════════════════════════════════════════════════════════

  function resetTimers() {
    clearTimeout(inactivityTimer);
    clearTimeout(warningTimer);

    // Programar cierre de sesión
    inactivityTimer = setTimeout(() => {
      cerrarSesion();
    }, INACTIVITY_TIMEOUT_MS);

    // Programar aviso de advertencia
    warningTimer = setTimeout(() => {
      mostrarAdvertencia();
    }, INACTIVITY_TIMEOUT_MS - WARNING_BEFORE_MS);
  }

  function mostrarAdvertencia() {
    // Evitar múltiples modales
    if (warningModal) return;

    const remainingSeconds = WARNING_BEFORE_MS / 1000;

    warningModal = document.createElement('div');
    warningModal.id = 'session-warning-modal';
    warningModal.innerHTML = `
      <div class="session-warning-backdrop"></div>
      <div class="session-warning-box">
        <div class="session-warning-icon">⏱️</div>
        <h3>Sesión a punto de expirar</h3>
        <p>Tu sesión se cerrará automáticamente en <strong id="session-countdown">${remainingSeconds}</strong> segundos por inactividad.</p>
        <div class="session-warning-actions">
          <button class="btn-keep-session" id="btn-keep-session">
            Mantener sesión activa
          </button>
        </div>
      </div>
    `;
    document.body.appendChild(warningModal);

    // Botón "Mantener sesión"
    const btnKeep = warningModal.querySelector('#btn-keep-session');
    btnKeep.addEventListener('click', () => {
      cerrarAdvertencia();
      resetTimers();
    });

    // Countdown dinámico
    let remaining = remainingSeconds;
    const countdownInterval = setInterval(() => {
      remaining--;
      const el = document.getElementById('session-countdown');
      if (el) el.textContent = remaining;
      if (remaining <= 0) {
        clearInterval(countdownInterval);
      }
    }, 1000);

    // Cerrar warning si el usuario hace algo
    warningModal.addEventListener('click', (e) => {
      if (e.target === warningModal || e.target.classList.contains('session-warning-backdrop')) {
        // No cerrar al hacer click fuera — forzar acción explícita
      }
    });
  }

  function cerrarAdvertencia() {
    if (warningModal) {
      warningModal.remove();
      warningModal = null;
    }
  }

  function cerrarSesion() {
    cerrarAdvertencia();

    // Limpiar almacenamiento según el rol
    try {
      sessionStorage.clear();
    } catch(e) {}

    // Mostrar mensaje breve antes de redirigir
    const overlay = document.createElement('div');
    overlay.id = 'session-expired-overlay';
    overlay.innerHTML = `
      <div class="session-expired-box">
        <div class="session-expired-icon">🔒</div>
        <h3>Sesión cerrada</h3>
        <p>Tu sesión se cerró por inactividad. Redirigiendo al login...</p>
      </div>
    `;
    document.body.appendChild(overlay);

    setTimeout(() => {
      window.location.href = 'index.html';
    }, 1500);
  }

  function actividadDetectada() {
    // Si hay modal de advertencia abierto, no resetear (usuario debe decidir)
    if (warningModal) return;
    resetTimers();
  }

  // ════════════════════════════════════════════════════════
  //  INICIALIZACIÓN
  // ════════════════════════════════════════════════════════

  function init() {
    // Escuchar eventos de actividad del usuario
    const eventos = [
      'mousedown', 'mousemove', 'keypress', 'scroll', 'touchstart',
      'click', 'wheel', 'resize', 'visibilitychange'
    ];

    eventos.forEach(evento => {
      document.addEventListener(evento, actividadDetectada, { passive: true });
    });

    // Iniciar timers
    resetTimers();

    // Guardar timestamp de inicio en sessionStorage para referencia
    sessionStorage.setItem(STORAGE_KEY, Date.now().toString());
  }

  // Ejecutar cuando el DOM esté listo
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();

// ════════════════════════════════════════════════════════
//  ESTILOS DINÁMICOS para modales
// ════════════════════════════════════════════════════════
const sessionStyles = document.createElement('style');
sessionStyles.textContent = `

/* ── MODAL DE ADVERTENCIA ── */
.session-warning-backdrop {
  position: fixed; inset: 0;
  background: rgba(0, 0, 0, 0.6);
  z-index: 99998;
  backdrop-filter: blur(4px);
}

.session-warning-box {
  position: fixed; top: 50%; left: 50%;
  transform: translate(-50%, -50%);
  background: white; border-radius: 20px;
  padding: 40px 32px; max-width: 400px; width: 90%;
  text-align: center; z-index: 99999;
  box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
  animation: slideUp 0.3s ease;
}

@keyframes slideUp {
  from { opacity: 0; transform: translate(-50%, -40%); }
  to   { opacity: 1; transform: translate(-50%, -50%); }
}

.session-warning-icon {
  font-size: 48px; margin-bottom: 16px;
}

.session-warning-box h3 {
  font-family: 'Playfair Display', serif;
  font-size: 22px; color: #1a1d2e; margin-bottom: 12px;
}

.session-warning-box p {
  font-size: 14px; color: #6b7280; line-height: 1.6;
  margin-bottom: 24px;
}

.session-warning-box #session-countdown {
  color: #dc2626; font-weight: 700;
}

.btn-keep-session {
  width: 100%; padding: 14px;
  background: #3b82f6; color: white;
  border: none; border-radius: 12px;
  font-size: 15px; font-weight: 600; cursor: pointer;
  transition: background 0.2s;
  font-family: 'DM Sans', sans-serif;
}

.btn-keep-session:hover {
  background: #1d4ed8;
}

/* ── OVERLAY SESIÓN EXPIRADA ── */
#session-expired-overlay {
  position: fixed; inset: 0;
  background: linear-gradient(135deg, #1e293b, #0f172a);
  display: flex; align-items: center; justify-content: center;
  z-index: 100000;
  animation: fadeIn 0.3s ease;
}

@keyframes fadeIn {
  from { opacity: 0; }
  to   { opacity: 1; }
}

.session-expired-box {
  text-align: center; color: white; padding: 40px;
}

.session-expired-icon {
  font-size: 64px; margin-bottom: 20px;
}

.session-expired-box h3 {
  font-family: 'Playfair Display', serif;
  font-size: 26px; margin-bottom: 12px;
}

.session-expired-box p {
  font-size: 14px; opacity: 0.7;
}
`;
document.head.appendChild(sessionStyles);
