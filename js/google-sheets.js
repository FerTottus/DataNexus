/**
 * Google Identity Services & Google Sheets API v4 Integration
 * Maneja el flujo OAuth 2.0 mediante GSI (Google Identity Services)
 * y realiza consultas directas a la API v4 de Google Sheets.
 */

const GoogleSheetsService = {
  // Scopes requeridos para leer hojas de cálculo de Google
  SCOPES: 'https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/userinfo.email',
  
  tokenClient: null,
  accessToken: null,
  tokenExpiresAt: null,
  userEmail: null,

  /**
   * Obtiene el Client ID almacenado o configurado por defecto
   */
  getClientId() {
    return localStorage.getItem('sheetpivot_google_client_id') || '';
  },

  /**
   * Guarda el Client ID en localStorage
   * @param {string} clientId 
   */
  setClientId(clientId) {
    localStorage.setItem('sheetpivot_google_client_id', (clientId || '').trim());
  },

  /**
   * Inicializa el cliente de token de Google Identity Services
   */
  initTokenClient(callback) {
    const clientId = this.getClientId();
    if (!clientId) {
      console.warn('Google Client ID no está configurado aún.');
      return false;
    }

    if (typeof google === 'undefined' || !google.accounts || !google.accounts.oauth2) {
      console.error('El SDK de Google Identity Services no está disponible.');
      return false;
    }

    try {
      this.tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: this.SCOPES,
        callback: async (resp) => {
          if (resp.error) {
            console.error('Error durante la autenticación OAuth:', resp);
            if (window.ClipboardUtil) {
              window.ClipboardUtil.showToast(`Error de autenticación: ${resp.error_description || resp.error}`, 'danger');
            }
            return;
          }

          this.accessToken = resp.access_token;
          // Estimar expiración (normalmente 3600 segundos)
          this.tokenExpiresAt = Date.now() + (resp.expires_in || 3500) * 1000;
          
          // Obtener email del usuario
          await this.fetchUserProfile();

          if (window.ClipboardUtil) {
            window.ClipboardUtil.showToast(`¡Conectado exitosamente con Google (${this.userEmail || 'OK'})!`, 'success');
          }

          if (callback) callback(true);
        },
      });
      return true;
    } catch (err) {
      console.error('Error al inicializar tokenClient de Google:', err);
      return false;
    }
  },

  /**
   * Solicita el token de acceso OAuth 2.0 (muestra la ventana emergente de Google)
   */
  requestAccessToken(callback) {
    const clientId = this.getClientId();
    if (!clientId) {
      if (window.ClipboardUtil) {
        window.ClipboardUtil.showToast('Por favor configura primero tu Google Client ID en el botón ⚙️', 'info', 4000);
      }
      return;
    }

    if (!this.tokenClient) {
      const initialized = this.initTokenClient(callback);
      if (!initialized) return;
    }

    // Solicitar token (abre el popup de selección de cuenta)
    this.tokenClient.requestAccessToken({ prompt: '' });
  },

  /**
   * Comprueba si el usuario está autenticado y el token sigue vigente
   */
  isAuthenticated() {
    return !!(this.accessToken && this.tokenExpiresAt && Date.now() < this.tokenExpiresAt);
  },

  /**
   * Cierra sesión localmente
   */
  logout() {
    if (this.accessToken && typeof google !== 'undefined' && google.accounts && google.accounts.oauth2) {
      try {
        google.accounts.oauth2.revoke(this.accessToken, () => {
          console.log('Token de acceso revocado.');
        });
      } catch (e) {
        console.warn('Error al revocar token:', e);
      }
    }
    this.accessToken = null;
    this.tokenExpiresAt = null;
    this.userEmail = null;
  },

  /**
   * Obtiene la información básica del usuario conectado
   */
  async fetchUserProfile() {
    if (!this.accessToken) return null;
    try {
      const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: {
          Authorization: `Bearer ${this.accessToken}`
        }
      });
      if (res.ok) {
        const data = await res.json();
        this.userEmail = data.email || null;
        return data;
      }
    } catch (e) {
      console.warn('No se pudo obtener el perfil de usuario:', e);
    }
    return null;
  },

  /**
   * Extrae el ID del Spreadsheet a partir de una URL completa o de un ID directo
   * @param {string} input - URL o ID del Sheet
   * @returns {string|null} ID extraído o null
   */
  extractSpreadsheetId(input) {
    if (!input) return null;
    const trimmed = input.trim();
    // Expresión regular para capturar el ID de Google Sheets
    const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      return match[1];
    }
    // Si no es URL, verificar si tiene formato de ID de Google Sheet (alfanumérico largo)
    if (/^[a-zA-Z0-9-_]{15,}$/.test(trimmed)) {
      return trimmed;
    }
    return null;
  },

  /**
   * Obtiene los metadatos del Spreadsheet (específicamente la lista de pestañas/hojas)
   * @param {string} spreadsheetId 
   * @returns {Promise<Array<{title: string, sheetId: number}>>}
   */
  async fetchSheetTabs(spreadsheetId) {
    if (!this.isAuthenticated()) {
      throw new Error('Debes iniciar sesión con Google para acceder a hojas de cálculo compartidas.');
    }

    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${this.accessToken}`
      }
    });

    if (!response.ok) {
      const errorJson = await response.json().catch(() => ({}));
      const message = errorJson.error?.message || `Error HTTP ${response.status}: No se pudo leer el archivo. Verifica que haya sido compartido con tu cuenta.`;
      throw new Error(message);
    }

    const data = await response.json();
    if (!data.sheets || !data.sheets.length) {
      throw new Error('El archivo no contiene hojas visibles.');
    }

    return data.sheets.map(s => ({
      title: s.properties.title,
      sheetId: s.properties.sheetId
    }));
  },

  /**
   * Lee todos los valores de una pestaña específica de la hoja de cálculo
   * @param {string} spreadsheetId 
   * @param {string} tabName 
   * @returns {Promise<{headers: Array<string>, rows: Array<Object>, rawValues: Array<Array<any>>}>}
   */
  async fetchSheetData(spreadsheetId, tabName) {
    if (!this.isAuthenticated()) {
      throw new Error('Debes iniciar sesión con Google para acceder al contenido.');
    }

    // Codificar el nombre de la pestaña para la URL
    const encodedRange = encodeURIComponent(tabName);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${this.accessToken}`
      }
    });

    if (!response.ok) {
      const errorJson = await response.json().catch(() => ({}));
      const message = errorJson.error?.message || `Error al obtener datos de la hoja "${tabName}".`;
      throw new Error(message);
    }

    const data = await response.json();
    const rawValues = data.values || [];

    if (rawValues.length === 0) {
      throw new Error(`La pestaña "${tabName}" está vacía.`);
    }

    return this.parseSheetMatrix(rawValues);
  },

  /**
   * Convierte una matriz 2D de valores de Sheets en encabezados y objetos de fila estructurados
   * @param {Array<Array<any>>} values 
   */
  parseSheetMatrix(values) {
    if (!values || values.length === 0) {
      return { headers: [], rows: [], rawValues: [] };
    }

    // La primera fila se asume como encabezados
    const rawHeaders = values[0];
    const headers = rawHeaders.map((h, idx) => {
      const title = h !== null && h !== undefined ? String(h).trim() : '';
      return title || `Columna_${idx + 1}`;
    });

    // Filas subsiguientes
    const rows = [];
    for (let r = 1; r < values.length; r++) {
      const rowArr = values[r];
      const rowObj = {};
      let hasData = false;

      for (let c = 0; c < headers.length; c++) {
        const header = headers[c];
        const val = rowArr[c] !== undefined ? rowArr[c] : null;
        rowObj[header] = val;
        if (val !== null && val !== '') {
          hasData = true;
        }
      }

      // Omitir filas totalmente vacías
      if (hasData) {
        rows.push(rowObj);
      }
    }

    return {
      headers,
      rows,
      rawValues: values
    };
  }
};

window.GoogleSheetsService = GoogleSheetsService;
