/**
 * DataNexus - Módulo de Dotación
 * Controlador de conexión con Google Drive / Sheets
 */

// Estado global del módulo Dotación
const DotacionState = {
  sheetId: '',
  sheetUrl: '',
  activeTab: '',
  availableTabs: [],
  dataset: {
    headers: [],
    rows: [],
    rawValues: []
  }
};

document.addEventListener('DOMContentLoaded', () => {
  initUIEvents();
  GoogleSheetsService.initAuth(); // Restaura token y correo si ya existían en localStorage
  checkAuthAndConfig();

  // Restaurar URL de hoja previa si existe en el almacenamiento local
  const savedUrl = localStorage.getItem('dotacion_sheet_url');
  if (savedUrl) {
    document.getElementById('sheetUrlInput').value = savedUrl;
  }
});

/**
 * Inicializa los eventos de botones y acciones del usuario
 */
function initUIEvents() {
  // 1. Botón Login de Google
  document.getElementById('btnGoogleLogin').addEventListener('click', () => {
    GoogleSheetsService.requestAccessToken((success) => {
      if (success) {
        updateAuthUI(true);
        if (document.getElementById('sheetUrlInput').value.trim()) {
          fetchSheetWorkbook();
        }
      }
    });
  });

  // 2. Botón Logout
  document.getElementById('btnGoogleLogout').addEventListener('click', () => {
    GoogleSheetsService.logout();
    updateAuthUI(false);
    if (window.ClipboardUtil) {
      ClipboardUtil.showToast('Sesión de Google cerrada', 'info');
    }
  });

  // 3. Botón Conectar y Cargar Sheet
  document.getElementById('btnFetchSheets').addEventListener('click', () => {
    fetchSheetWorkbook();
  });

  // 4. Botón Explorar Drive
  document.getElementById('btnBrowseDrive').addEventListener('click', () => {
    if (!GoogleSheetsService.isAuthenticated()) {
      GoogleSheetsService.requestAccessToken((success) => {
        if (success) {
          updateAuthUI(true);
          openDriveModal();
        }
      });
    } else {
      openDriveModal();
    }
  });

  // 5. Cerrar modal de Drive
  document.getElementById('close-modal-btn').addEventListener('click', () => {
    document.getElementById('drive-modal').classList.add('hidden');
  });

  // Cerrar modal al hacer clic en el backdrop exterior
  document.getElementById('drive-modal').addEventListener('click', (e) => {
    if (e.target.id === 'drive-modal') {
      e.target.classList.add('hidden');
    }
  });

  // 6. Enter en el campo de texto de URL
  document.getElementById('sheetUrlInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      fetchSheetWorkbook();
    }
  });

  // 7. Selector de pestaña / hoja
  document.getElementById('sheetTabSelect').addEventListener('change', (e) => {
    const selectedTab = e.target.value;
    if (selectedTab) {
      loadTabData(selectedTab);
    }
  });
}

/**
 * Comprueba el estado de autenticación y actualiza la UI
 */
function checkAuthAndConfig() {
  const isAuth = GoogleSheetsService.isAuthenticated();
  updateAuthUI(isAuth);
}

/**
 * Actualiza los elementos visuales de login / usuario conectado
 */
function updateAuthUI(isAuthenticated) {
  const btnLogin = document.getElementById('btnGoogleLogin');
  const userInfo = document.getElementById('userInfo');
  const userEmailSpan = document.getElementById('userEmail');

  if (isAuthenticated) {
    btnLogin.classList.add('hidden');
    userInfo.classList.remove('hidden');
    userEmailSpan.innerText = GoogleSheetsService.userEmail || 'Google Conectado';
  } else {
    btnLogin.classList.remove('hidden');
    userInfo.classList.add('hidden');
  }
}

/**
 * Abre el modal y carga los archivos de Google Sheets recientes desde Drive
 */
async function openDriveModal() {
  const modal = document.getElementById('drive-modal');
  const loading = document.getElementById('drive-loading');
  const empty = document.getElementById('drive-empty');
  const list = document.getElementById('drive-file-list');
  const searchInput = document.getElementById('drive-search-input');

  modal.classList.remove('hidden');
  loading.classList.remove('hidden');
  list.innerHTML = '';
  empty.classList.add('hidden');
  if (searchInput) searchInput.value = '';

  const files = await GoogleSheetsService.fetchRecentSpreadsheets();
  loading.classList.add('hidden');

  if (!files || files.length === 0) {
    empty.classList.remove('hidden');
    return;
  }

  files.forEach(file => {
    const li = document.createElement('li');
    li.className = 'file-list-item';
    li.dataset.filename = file.name.toLowerCase();

    const date = new Date(file.modifiedTime).toLocaleDateString('es-ES', {
      day: 'numeric',
      month: 'short',
      year: 'numeric'
    });

    let isShared = true;
    let ownerName = 'Desconocido';

    if (file.owners && file.owners.length > 0) {
      if (file.owners[0].me) {
        isShared = false;
        ownerName = 'Tú';
      } else {
        ownerName = file.owners[0].displayName || 'Compartido';
      }
    }

    const isExcel = file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.name.toLowerCase().endsWith('.xlsx');
    const badgeClass = isShared ? 'badge-shared' : 'badge-owned';
    const badgeText = isShared ? 'Compartido' : 'Mío';
    const iconClass = isExcel ? 'fa-triangle-exclamation text-danger' : (isShared ? 'fa-file-excel file-icon-shared' : 'fa-file-excel file-icon-owned');
    const ownerIcon = isShared ? 'fa-users' : 'fa-user';

    const excelWarningHtml = isExcel ? `<span class="file-badge custom-tooltip-icon" style="background: #fee2e2; color: #ef4444; border: 1px solid #fecaca; margin-right: 5px;" data-tooltip="Ábrelo en Drive y dale a 'Guardar como Hoja de cálculo de Google'">⚠️ Inválido (.xlsx)</span>` : '';

    li.innerHTML = `
      <div class="file-item-icon">
        <i class="fa-solid ${iconClass}"></i>
      </div>
      <div class="file-item-content">
        <div class="file-item-title" title="${file.name}">${file.name}</div>
        <div class="file-item-meta">
          <span class="owner"><i class="fa-solid ${ownerIcon}"></i> ${ownerName}</span>
          <span class="date"><i class="fa-regular fa-clock"></i> ${date}</span>
        </div>
      </div>
      <div class="file-item-action">
        ${excelWarningHtml}
        <span class="file-badge ${badgeClass}">${badgeText}</span>
      </div>
    `;

    li.addEventListener('click', () => {
      document.getElementById('sheetUrlInput').value = file.id;
      modal.classList.add('hidden');
      fetchSheetWorkbook();
    });

    list.appendChild(li);
  });
}

/**
 * Filtra la lista de archivos mostrados en el modal de Drive
 */
window.filterDriveList = function(searchTerm) {
  const term = searchTerm.toLowerCase();
  const items = document.querySelectorAll('#drive-file-list .file-list-item');
  let hasVisible = false;

  items.forEach(item => {
    if (item.dataset.filename.includes(term)) {
      item.style.display = 'flex';
      hasVisible = true;
    } else {
      item.style.display = 'none';
    }
  });

  const empty = document.getElementById('drive-empty');
  if (hasVisible) {
    empty.classList.add('hidden');
  } else {
    empty.classList.remove('hidden');
  }
};

/**
 * Carga el libro de Google Sheets y sus pestañas
 */
async function fetchSheetWorkbook() {
  const input = document.getElementById('sheetUrlInput').value.trim();
  if (!input) {
    if (window.ClipboardUtil) {
      ClipboardUtil.showToast('Ingresa la URL o ID del Google Sheet compartido', 'info');
    }
    return;
  }

  const sheetId = GoogleSheetsService.extractSpreadsheetId(input);
  if (!sheetId) {
    if (window.ClipboardUtil) {
      ClipboardUtil.showToast('URL o ID de Google Sheet no válido', 'danger');
    }
    return;
  }

  if (!GoogleSheetsService.isAuthenticated()) {
    if (window.ClipboardUtil) {
      ClipboardUtil.showToast('Inicia sesión con Google para acceder al Sheet', 'info');
    }
    await GoogleSheetsService.requestAccessToken(async (success) => {
      if (success) {
        updateAuthUI(true);
        await performFetchTabs(sheetId, input);
      }
    });
    return;
  }

  await performFetchTabs(sheetId, input);
}

/**
 * Realiza la llamada a la API para obtener las pestañas del documento
 */
async function performFetchTabs(sheetId, originalInput) {
  const badge = document.getElementById('sheetStatusBadge');
  badge.className = 'badge badge-warning';
  badge.innerText = 'Cargando...';

  try {
    const tabs = await GoogleSheetsService.fetchSheetTabs(sheetId);

    DotacionState.sheetId = sheetId;
    DotacionState.sheetUrl = originalInput;
    DotacionState.availableTabs = tabs;

    // Guardar para conveniencia del usuario en próximas visitas
    localStorage.setItem('dotacion_sheet_url', originalInput);

    // Poblar selector de pestañas
    const select = document.getElementById('sheetTabSelect');
    select.innerHTML = '';
    tabs.forEach(tab => {
      const opt = document.createElement('option');
      opt.value = tab.title;
      opt.innerText = tab.title;
      select.appendChild(opt);
    });

    document.getElementById('tabSelectorContainer').classList.remove('hidden');

    // Cargar la primera pestaña por defecto
    const tabToLoad = tabs[0].title;
    select.value = tabToLoad;
    await loadTabData(tabToLoad);

    badge.className = 'badge badge-success';
    badge.innerText = 'Conectado';
  } catch (err) {
    console.error(err);
    badge.className = 'badge badge-warning';
    badge.innerText = 'Error';
    if (window.ClipboardUtil) {
      ClipboardUtil.showToast(err.message || 'Error al conectar con Google Sheet', 'danger', 5000);
    }
  }
}

/**
 * Carga los datos de una pestaña seleccionada
 */
async function loadTabData(tabName) {
  try {
    DotacionState.activeTab = tabName;
    const parsedData = await GoogleSheetsService.fetchSheetData(DotacionState.sheetId, tabName);
    DotacionState.dataset = parsedData;

    updateDatasetInfoBar();

    if (window.ClipboardUtil) {
      ClipboardUtil.showToast(`Pestaña "${tabName}" cargada exitosamente (${parsedData.rows.length} filas)`, 'success');
    }
  } catch (err) {
    console.error(err);
    if (window.ClipboardUtil) {
      ClipboardUtil.showToast(err.message || 'Error al leer datos de la pestaña', 'danger');
    }
  }
}

/**
 * Actualiza la barra con la metadata de filas, columnas y hora
 */
function updateDatasetInfoBar() {
  const container = document.getElementById('datasetInfo');
  container.classList.remove('hidden');

  document.getElementById('infoSheetName').innerText = DotacionState.activeTab || '-';
  document.getElementById('infoRowCount').innerText = DotacionState.dataset.rows.length.toLocaleString();
  document.getElementById('infoColCount').innerText = DotacionState.dataset.headers.length;
  document.getElementById('infoTimestamp').innerText = new Date().toLocaleTimeString();
}
