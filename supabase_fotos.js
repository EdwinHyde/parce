// ════════════════════════════════════════════════════════
//  PARCE — Módulo de Fotos de Perfil (Supabase Storage)
//  v1.0
// ════════════════════════════════════════════════════════
//
//  INSTRUCCIONES DE CONFIGURACIÓN:
//  1. Crear un proyecto en https://supabase.com
//  2. En el dashboard del proyecto, ir a Storage
//  3. Crear un bucket llamado "fotos-perfil" con acceso público
//  4. En las políticas del bucket, agregar:
//     - SELECT (download): permitir a todos (anon)
//     - INSERT (upload):   permitir a todos (anon)
//     - UPDATE (update):   permitir a todos (anon)
//     - DELETE (delete):   permitir a todos (anon)
//  5. Copiar la URL del proyecto y la anon key pública
//  6. Pegar los valores en SUPABASE_URL y SUPABASE_ANON_KEY abajo
//
// ════════════════════════════════════════════════════════

// ══ CONFIGURAR ESTOS VALORES CON TU PROYECTO SUPABASE ══
const SUPABASE_URL      = 'https://sccrnbhuyzdfmnaivjna.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_R0SkPEskfvWakmqu4kNmmQ_HVbTJU8m';
const SUPABASE_BUCKET   = 'fotos-perfil';

// ════════════════════════════════════════════════════════
//  Funciones de Supabase Storage
// ════════════════════════════════════════════════════════

/**
 * Sube una foto de perfil a Supabase Storage
 * @param {File} file - Archivo de imagen seleccionado
 * @param {string} tipoUsuario - 'docente' | 'estudiante' | 'orientador'
 * @param {string} identificador - ID o documento del usuario (se usa como nombre de archivo)
 * @returns {Promise<{url: string, error: string|null}>}
 */
async function subirFotoPerfil(file, tipoUsuario, identificador) {
  try {
    // Validar archivo
    if (!file) return { url: null, error: 'No se seleccionó ningún archivo.' };

    const tiposPermitidos = ['image/jpeg', 'image/png', 'image/webp'];
    if (!tiposPermitidos.includes(file.type)) {
      return { url: null, error: 'Solo se permiten imágenes JPG, PNG o WEBP.' };
    }

    const maxSize = 2 * 1024 * 1024; // 2MB
    if (file.size > maxSize) {
      return { url: null, error: 'La imagen no debe superar los 2MB.' };
    }

    // Redimensionar imagen antes de subir (max 400x400)
    const resizedBlob = await redimensionarImagen(file, 400);

    // Generar nombre de archivo único
    const extension = file.name.split('.').pop().toLowerCase() || 'jpg';
    const nombreArchivo = `${tipoUsuario}/${limpiarNombreArchivo(identificador)}.${extension}`;

    // Subir a Supabase Storage
    const response = await fetch(
      `${SUPABASE_URL}/storage/v1/object/${SUPABASE_BUCKET}/${nombreArchivo}`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
          'apikey': SUPABASE_ANON_KEY,
          'Content-Type': file.type,
          'x-upsert': 'true' // Sobreescribir si ya existe
        },
        body: resizedBlob
      }
    );

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      return { url: null, error: errorData.message || 'Error al subir la imagen.' };
    }

    // Construir URL pública
    const urlPublica = `${SUPABASE_URL}/storage/v1/object/public/${SUPABASE_BUCKET}/${nombreArchivo}`;

    return { url: urlPublica, error: null };

  } catch (err) {
    console.error('Error subiendo foto:', err);
    return { url: null, error: 'Error de conexión al subir la imagen.' };
  }
}

/**
 * Elimina una foto de perfil de Supabase Storage
 * @param {string} tipoUsuario - 'docente' | 'estudiante' | 'orientador'
 * @param {string} identificador - ID del usuario
 * @returns {Promise<{ok: boolean, error: string|null}>}
 */
async function eliminarFotoPerfil(tipoUsuario, identificador) {
  try {
    // Intentar eliminar con extensiones comunes
    const extensiones = ['jpg', 'jpeg', 'png', 'webp'];
    for (const ext of extensiones) {
      const nombreArchivo = `${tipoUsuario}/${limpiarNombreArchivo(identificador)}.${ext}`;
      await fetch(
        `${SUPABASE_URL}/storage/v1/object/${SUPABASE_BUCKET}/${nombreArchivo}`,
        {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
            'apikey': SUPABASE_ANON_KEY
          }
        }
      );
    }
    return { ok: true, error: null };
  } catch (err) {
    return { ok: false, error: 'Error al eliminar la foto.' };
  }
}

/**
 * Obtiene la URL pública de la foto de perfil desde la columna FOTO_URL
 * Si no existe en la base de datos, retorna null
 * @param {string} fotoUrl - URL almacenada en la hoja de cálculo
 * @returns {string|null}
 */
function obtenerUrlFoto(fotoUrl) {
  if (!fotoUrl || fotoUrl.trim() === '' || fotoUrl === '—') return null;
  return fotoUrl.trim();
}

// ════════════════════════════════════════════════════════
//  UTILIDADES INTERNAS
// ════════════════════════════════════════════════════════

/**
 * Redimensiona una imagen manteniendo proporción
 * @param {File} file
 * @param {number} maxDim - Dimensión máxima (ancho o alto)
 * @returns {Promise<Blob>}
 */
function redimensionarImagen(file, maxDim) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      const img = new Image();
      img.onload = function() {
        let w = img.width, h = img.height;
        if (w > maxDim || h > maxDim) {
          if (w > h) { h = Math.round(h * maxDim / w); w = maxDim; }
          else       { w = Math.round(w * maxDim / h); h = maxDim; }
        }

        const canvas = document.createElement('canvas');
        canvas.width  = w;
        canvas.height = h;
        const ctx = canvas.getContext('2d');

        // Fondo blanco para transparencias
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);

        canvas.toBlob(blob => {
          if (blob) resolve(blob);
          else reject(new Error('Error al procesar la imagen.'));
        }, 'image/jpeg', 0.85);
      };
      img.onerror = () => reject(new Error('No se pudo leer la imagen.'));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error('Error leyendo el archivo.'));
    reader.readAsDataURL(file);
  });
}

/**
 * Limpia el identificador para usarlo como nombre de archivo
 */
function limpiarNombreArchivo(str) {
  return String(str || 'sin_id').trim()
    .replace(/\s+/g, '_')
    .replace(/[^a-zA-Z0-9_\-]/g, '')
    .substring(0, 60) || 'sin_id';
}

// ════════════════════════════════════════════════════════
//  COMPONENTE UI — Widget de foto de perfil
//  Genera el HTML y lógica para subir/ver foto en el perfil
// ════════════════════════════════════════════════════════

/**
 * Crea el widget HTML de foto de perfil
 * @param {Object} opciones
 * @param {string} opciones.containerId - ID del elemento donde insertar el widget
 * @param {string} opciones.fotoUrl - URL actual de la foto (o null)
 * @param {string} opciones.iniciales - Letra(s) para el avatar por defecto
 * @param {string} opciones.tipoUsuario - 'docente' | 'estudiante'
 * @param {string} opciones.identificador - ID/documento del usuario
 * @param {Function} opciones.onSubida - Callback cuando se sube foto: (url) => {}
 */
function crearWidgetFotoPerfil(opciones) {
  const { containerId, fotoUrl, iniciales, tipoUsuario, identificador, onSubida } = opciones;
  const container = document.getElementById(containerId);
  if (!container) return;

  const tieneSupabase = SUPABASE_URL !== 'https://TU_PROYECTO.supabase.co';

  const avatarHTML = fotoUrl
    ? `<img src="${fotoUrl}" alt="Foto de perfil" class="foto-perfil-img" id="foto-perfil-preview">`
    : `<div class="foto-perfil-iniciales" id="foto-perfil-preview">${iniciales}</div>`;

  container.innerHTML = `
    <div class="foto-perfil-widget">
      <div class="foto-perfil-avatar-wrap">
        ${avatarHTML}
        ${tieneSupabase ? `
          <label class="foto-perfil-cambiar" for="foto-perfil-input" title="Cambiar foto">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/>
              <circle cx="12" cy="13" r="4"/>
            </svg>
          </label>
          <input type="file" id="foto-perfil-input" accept="image/jpeg,image/png,image/webp" style="display:none;">
        ` : ''}
      </div>
      <div id="foto-perfil-status" class="foto-perfil-status" style="display:none;"></div>
    </div>
  `;

  if (!tieneSupabase) return;

  // Listener para subida
  const fileInput = document.getElementById('foto-perfil-input');
  fileInput.addEventListener('change', async function() {
    const file = this.files[0];
    if (!file) return;

    const status = document.getElementById('foto-perfil-status');
    status.style.display = 'block';
    status.textContent = 'Subiendo foto...';
    status.className = 'foto-perfil-status foto-status-loading';

    const result = await subirFotoPerfil(file, tipoUsuario, identificador);

    if (result.error) {
      status.textContent = result.error;
      status.className = 'foto-perfil-status foto-status-error';
      setTimeout(() => { status.style.display = 'none'; }, 4000);
    } else {
      // Actualizar preview
      const preview = document.getElementById('foto-perfil-preview');
      if (preview.tagName === 'IMG') {
        preview.src = result.url + '?t=' + Date.now();
      } else {
        preview.outerHTML = `<img src="${result.url}?t=${Date.now()}" alt="Foto de perfil" class="foto-perfil-img" id="foto-perfil-preview">`;
      }

      status.textContent = '✓ Foto actualizada';
      status.className = 'foto-perfil-status foto-status-ok';
      setTimeout(() => { status.style.display = 'none'; }, 3000);

      // Callback para guardar URL en la base de datos
      if (onSubida) onSubida(result.url);
    }

    // Limpiar input
    this.value = '';
  });
}
