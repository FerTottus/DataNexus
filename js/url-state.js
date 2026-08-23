/**
 * URL State Manager for SheetPivot
 * Gestiona la persistencia y compresión del estado de la aplicación en el URL hash
 * usando LZ-String (URL-Safe encoding).
 */

const UrlStateManager = {
  HASH_PREFIX: '#state=',

  /**
   * Comprime y serializa un objeto de estado completo
   * @param {Object} stateObj 
   * @returns {string} Hash comprimido con prefijo
   */
  encodeState(stateObj) {
    try {
      const jsonString = JSON.stringify(stateObj);
      if (typeof LZString !== 'undefined') {
        const compressed = LZString.compressToEncodedURIComponent(jsonString);
        return this.HASH_PREFIX + compressed;
      } else {
        // Fallback Base64 si LZString no estuviese disponible
        return this.HASH_PREFIX + encodeURIComponent(btoa(unescape(encodeURIComponent(jsonString))));
      }
    } catch (err) {
      console.error('Error al codificar el estado en la URL:', err);
      return '';
    }
  },

  /**
   * Lee y descomprime el estado desde la URL actual
   * @returns {Object|null} Objeto de estado restaurado o null
   */
  decodeState() {
    try {
      const hash = window.location.hash;
      if (!hash || !hash.startsWith(this.HASH_PREFIX)) {
        return null;
      }

      const payload = hash.substring(this.HASH_PREFIX.length);
      if (!payload) return null;

      let jsonString = '';
      if (typeof LZString !== 'undefined') {
        jsonString = LZString.decompressFromEncodedURIComponent(payload);
      }

      if (!jsonString) {
        // Intentar fallback
        try {
          jsonString = decodeURIComponent(escape(atob(decodeURIComponent(payload))));
        } catch (e) {
          console.warn('Fallback decompression falló:', e);
        }
      }

      if (!jsonString) return null;
      return JSON.parse(jsonString);
    } catch (err) {
      console.error('Error al decodificar estado de la URL:', err);
      return null;
    }
  },

  /**
   * Actualiza el hash de la URL en el navegador sin recargar la página
   * @param {Object} stateObj 
   */
  persistToUrl(stateObj) {
    const hash = this.encodeState(stateObj);
    if (hash) {
      // Usar replaceState para no llenar el historial con cada cambio menor
      history.replaceState(null, '', hash);
    }
  },

  /**
   * Obtiene la URL completa compartible con el estado actual
   * @param {Object} stateObj 
   * @returns {string} URL absoluta lista para compartir
   */
  getShareableUrl(stateObj) {
    const hash = this.encodeState(stateObj);
    const baseUrl = window.location.origin + window.location.pathname;
    return baseUrl + hash;
  },

  /**
   * Copia la URL compartible al portapapeles y notifica al usuario
   * @param {Object} stateObj 
   */
  async copyShareableUrl(stateObj) {
    const url = this.getShareableUrl(stateObj);
    try {
      await navigator.clipboard.writeText(url);
      if (window.ClipboardUtil) {
        window.ClipboardUtil.showToast('🔗 ¡Enlace comprimido copiado al portapapeles!', 'success');
      }
    } catch (err) {
      console.error('Error al copiar URL compartible:', err);
      if (window.ClipboardUtil) {
        window.ClipboardUtil.showToast('No se pudo copiar el enlace', 'danger');
      }
    }
  }
};

window.UrlStateManager = UrlStateManager;
