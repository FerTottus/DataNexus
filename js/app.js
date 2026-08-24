/**
 * SheetPivot Studio - Main Application Controller
 * Coordina la autenticación, carga de Google Sheets, gestión de múltiples tablas,
 * filtros, copiado al portapapeles y sincronización de URL con compresión LZ-String.
 */

// Estado global de la aplicación
const AppState = {
  sheetId: '',
  sheetUrl: '',
  activeTab: '',
  availableTabs: [],
  dataset: {
    headers: [],
    rows: [],
    rawValues: []
  },
  tables: [], // Lista de tablas creadas por el usuario
  isDemo: false
};

// Conjunto de datos de demostración para pruebas inmediatas
const DEMO_DATASET = {
  headers: ['Fecha', 'Región', 'Vendedor', 'Categoría', 'Producto', 'Cantidad', 'Precio_Unitario', 'Ventas_Totales', 'Estado'],
  rows: [
    { Fecha: '2026-01-10', Región: 'Norte', Vendedor: 'Carlos Ruiz', Categoría: 'Electrónica', Producto: 'Laptop Pro 15', Cantidad: 4, Precio_Unitario: 1200, Ventas_Totales: 4800, Estado: 'Completado' },
    { Fecha: '2026-01-12', Región: 'Sur', Vendedor: 'Ana Gómez', Categoría: 'Muebles', Producto: 'Silla Ergonómica', Cantidad: 12, Precio_Unitario: 150, Ventas_Totales: 1800, Estado: 'Completado' },
    { Fecha: '2026-01-15', Región: 'Centro', Vendedor: 'Luis Morales', Categoría: 'Electrónica', Producto: 'Monitor 4K 27"', Cantidad: 8, Precio_Unitario: 350, Ventas_Totales: 2800, Estado: 'Completado' },
    { Fecha: '2026-01-18', Región: 'Norte', Vendedor: 'Carlos Ruiz', Categoría: 'Accesorios', Producto: 'Teclado Mecánico', Cantidad: 20, Precio_Unitario: 80, Ventas_Totales: 1600, Estado: 'Pendiente' },
    { Fecha: '2026-01-20', Región: 'Sur', Vendedor: 'Ana Gómez', Categoría: 'Electrónica', Producto: 'Laptop Pro 15', Cantidad: 6, Precio_Unitario: 1200, Ventas_Totales: 7200, Estado: 'Completado' },
    { Fecha: '2026-01-22', Región: 'Este', Vendedor: 'Sofía Castro', Categoría: 'Muebles', Producto: 'Escritorio Elevable', Cantidad: 5, Precio_Unitario: 400, Ventas_Totales: 2000, Estado: 'Completado' },
    { Fecha: '2026-01-25', Región: 'Centro', Vendedor: 'Luis Morales', Categoría: 'Accesorios', Producto: 'Mouse Inalámbrico', Cantidad: 30, Precio_Unitario: 45, Ventas_Totales: 1350, Estado: 'Completado' },
    { Fecha: '2026-01-28', Región: 'Este', Vendedor: 'Sofía Castro', Categoría: 'Electrónica', Producto: 'Monitor 4K 27"', Cantidad: 10, Precio_Unitario: 350, Ventas_Totales: 3500, Estado: 'Completado' },
    { Fecha: '2026-02-02', Región: 'Norte', Vendedor: 'Carlos Ruiz', Categoría: 'Muebles', Producto: 'Silla Ergonómica', Cantidad: 15, Precio_Unitario: 150, Ventas_Totales: 2250, Estado: 'Completado' },
    { Fecha: '2026-02-05', Región: 'Sur', Vendedor: 'Ana Gómez', Categoría: 'Accesorios', Producto: 'Teclado Mecánico', Cantidad: 14, Precio_Unitario: 80, Ventas_Totales: 1120, Estado: 'Cancelado' },
    { Fecha: '2026-02-08', Región: 'Centro', Vendedor: 'Luis Morales', Categoría: 'Electrónica', Producto: 'Laptop Pro 15', Cantidad: 3, Precio_Unitario: 1200, Ventas_Totales: 3600, Estado: 'Completado' },
    { Fecha: '2026-02-12', Región: 'Este', Vendedor: 'Sofía Castro', Categoría: 'Muebles', Producto: 'Silla Ergonómica', Cantidad: 8, Precio_Unitario: 150, Ventas_Totales: 1200, Estado: 'Pendiente' }
  ]
};

document.addEventListener('DOMContentLoaded', () => {
  initUIEvents();
  checkAuthAndConfig();
  restoreStateFromUrl();
});

/**
 * Inicializa los eventos de botones
 */
function initUIEvents() {
  // 1. Botón Login de Google
  document.getElementById('btnGoogleLogin').addEventListener('click', () => {
    GoogleSheetsService.requestAccessToken((success) => {
      if (success) {
        updateAuthUI(true);
        // Si ya había un sheet ingresado, intentar cargarlo
        if (AppState.sheetUrl) {
          fetchSheetWorkbook();
        }
      }
    });
  });

  // 2. Botón Logout
  document.getElementById('btnGoogleLogout').addEventListener('click', () => {
    GoogleSheetsService.logout();
    updateAuthUI(false);
    ClipboardUtil.showToast('Sesión de Google cerrada', 'info');
  });

  // 3. Botón de Conectar y Cargar Google Sheet
  document.getElementById('btnFetchSheets').addEventListener('click', () => {
    fetchSheetWorkbook();
  });

  // Entrada con tecla Enter en el input de URL
  document.getElementById('sheetUrlInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      fetchSheetWorkbook();
    }
  });

  // 4. Selector de Pestaña de Hoja
  document.getElementById('sheetTabSelect').addEventListener('change', (e) => {
    const selectedTab = e.target.value;
    if (selectedTab) {
      loadTabData(selectedTab);
    }
  });

  // 5. Botón Cargar Datos Demo
  document.getElementById('btnDemoData').addEventListener('click', () => {
    loadDemoDataset();
  });

  // 6. Botón Compartir / Guardar URL
  document.getElementById('btnShareUrl').addEventListener('click', () => {
    syncStateToUrl(true);
  });

  // 7. Botones para agregar tablas
  document.getElementById('btnAddPivotTable').addEventListener('click', () => {
    addNewTable('pivot');
  });

  document.getElementById('btnAddFlatTable').addEventListener('click', () => {
    addNewTable('flat');
  });

  // 8. Limpiar espacio de trabajo
  document.getElementById('btnResetWorkspace').addEventListener('click', () => {
    if (confirm('¿Estás seguro de que deseas eliminar todas las tablas creadas?')) {
      AppState.tables = [];
      renderAllTables();
      syncStateToUrl();
      ClipboardUtil.showToast('Espacio de trabajo limpiado', 'info');
    }
  });
}

/**
 * Verifica si el usuario está autenticado
 */
function checkAuthAndConfig() {
  const isAuth = GoogleSheetsService.isAuthenticated();
  updateAuthUI(isAuth);
}

/**
 * Actualiza los elementos visuales de login/usuario
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
 * Carga la información del libro de Google Sheets y sus pestañas
 */
async function fetchSheetWorkbook() {
  const input = document.getElementById('sheetUrlInput').value.trim();
  if (!input) {
    ClipboardUtil.showToast('Ingresa la URL o ID del Google Sheet compartido', 'info');
    return;
  }

  const sheetId = GoogleSheetsService.extractSpreadsheetId(input);
  if (!sheetId) {
    ClipboardUtil.showToast('URL o ID de Google Sheet no válido', 'danger');
    return;
  }

  if (!GoogleSheetsService.isAuthenticated()) {
    ClipboardUtil.showToast('Inicia sesión con Google para acceder al Sheet', 'info');
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

async function performFetchTabs(sheetId, originalInput) {
  const badge = document.getElementById('sheetStatusBadge');
  badge.className = 'badge badge-warning';
  badge.innerText = 'Cargando...';

  try {
    const tabs = await GoogleSheetsService.fetchSheetTabs(sheetId);
    AppState.sheetId = sheetId;
    AppState.sheetUrl = originalInput;
    AppState.availableTabs = tabs;
    AppState.isDemo = false;

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

    // Cargar la primera pestaña por defecto o la que ya esté en el estado
    const tabToLoad = AppState.activeTab && tabs.some(t => t.title === AppState.activeTab)
      ? AppState.activeTab
      : tabs[0].title;

    select.value = tabToLoad;
    await loadTabData(tabToLoad);

    badge.className = 'badge badge-success';
    badge.innerText = 'Conectado';
  } catch (err) {
    console.error(err);
    badge.className = 'badge badge-warning';
    badge.innerText = 'Error';
    ClipboardUtil.showToast(err.message || 'Error al conectar con Google Sheet', 'danger', 5000);
  }
}

/**
 * Carga los datos de una pestaña seleccionada
 */
async function loadTabData(tabName) {
  try {
    AppState.activeTab = tabName;
    const parsedData = await GoogleSheetsService.fetchSheetData(AppState.sheetId, tabName);
    AppState.dataset = parsedData;

    updateDatasetInfoBar();
    renderAllTables();
    syncStateToUrl();

    ClipboardUtil.showToast(`Pestaña "${tabName}" cargada exitosamente (${parsedData.rows.length} filas)`, 'success');
  } catch (err) {
    console.error(err);
    ClipboardUtil.showToast(err.message || 'Error al leer datos de la pestaña', 'danger');
  }
}

/**
 * Carga el dataset de demostración
 */
function loadDemoDataset() {
  AppState.sheetId = 'demo-sheet-id';
  AppState.sheetUrl = 'https://docs.google.com/spreadsheets/d/demo-ventas-ejemplo/edit';
  AppState.activeTab = 'Ventas_2026';
  AppState.availableTabs = [{ title: 'Ventas_2026', sheetId: 0 }];
  AppState.dataset = DEMO_DATASET;
  AppState.isDemo = true;

  document.getElementById('sheetUrlInput').value = AppState.sheetUrl;
  
  const select = document.getElementById('sheetTabSelect');
  select.innerHTML = '<option value="Ventas_2026">Ventas_2026 (Demo)</option>';
  document.getElementById('tabSelectorContainer').classList.remove('hidden');

  const badge = document.getElementById('sheetStatusBadge');
  badge.className = 'badge badge-info';
  badge.innerText = 'Modo Demo';

  updateDatasetInfoBar();

  // Si no hay tablas, crear dos tablas de ejemplo
  if (AppState.tables.length === 0) {
    // 1. Tabla Dinámica de Ventas por Región y Categoría
    AppState.tables.push({
      id: 'table_' + Date.now(),
      type: 'pivot',
      title: 'Ventas Totales por Región y Categoría',
      rowFields: ['Región'],
      colFields: ['Categoría'],
      valField: 'Ventas_Totales',
      aggFunc: 'SUM',
      filters: [],
      globalSearch: '',
      isCollapsed: false
    });

    // 2. Tabla Plana de Registros Filtrados
    AppState.tables.push({
      id: 'table_' + (Date.now() + 1),
      type: 'flat',
      title: 'Detalle de Pedidos Completados',
      selectedColumns: ['Fecha', 'Vendedor', 'Producto', 'Cantidad', 'Ventas_Totales', 'Estado'],
      filters: [{ field: 'Estado', operator: 'equals', value: 'Completado' }],
      globalSearch: '',
      sortBy: 'Ventas_Totales',
      sortOrder: 'desc',
      limit: 100,
      isCollapsed: false
    });
  }

  renderAllTables();
  syncStateToUrl();
  ClipboardUtil.showToast('Datos de demostración cargados. ¡Prueba a copiar y modificar las tablas!', 'success');
}

/**
 * Actualiza la barra con la metadata de filas, columnas y hora
 */
function updateDatasetInfoBar() {
  const container = document.getElementById('datasetInfo');
  container.classList.remove('hidden');

  document.getElementById('infoSheetName').innerText = AppState.activeTab || '-';
  document.getElementById('infoRowCount').innerText = AppState.dataset.rows.length.toLocaleString();
  document.getElementById('infoColCount').innerText = AppState.dataset.headers.length;
  document.getElementById('infoTimestamp').innerText = new Date().toLocaleTimeString();
}

/**
 * Agrega una nueva tabla al espacio de trabajo
 * @param {'pivot'|'flat'} type 
 */
function addNewTable(type = 'pivot') {
  if (!AppState.dataset.headers || AppState.dataset.headers.length === 0) {
    ClipboardUtil.showToast('Primero conecta una hoja de cálculo o carga los datos de demostración.', 'info');
    return;
  }

  const headers = AppState.dataset.headers;
  const tableId = 'table_' + Date.now() + '_' + Math.floor(Math.random() * 1000);

  if (type === 'pivot') {
    // Buscar campo numérico sugerido para valor
    const numericCandidate = headers.find(h => {
      const sample = AppState.dataset.rows[0]?.[h];
      return typeof sample === 'number' || (!isNaN(Number(sample)) && sample !== null && sample !== '');
    }) || headers[headers.length - 1];

    AppState.tables.push({
      id: tableId,
      type: 'pivot',
      title: `Tabla Dinámica ${AppState.tables.length + 1}`,
      rowFields: [headers[0]],
      colFields: headers.length > 1 ? [headers[1]] : [],
      valField: numericCandidate || headers[0],
      aggFunc: 'SUM',
      filters: [],
      globalSearch: '',
      isCollapsed: false
    });
  } else {
    // Tabla plana
    AppState.tables.push({
      id: tableId,
      type: 'flat',
      title: `Vista Filtrada ${AppState.tables.length + 1}`,
      selectedColumns: [...headers],
      filters: [],
      globalSearch: '',
      sortBy: '',
      sortOrder: 'asc',
      limit: 100,
      isCollapsed: false
    });
  }

  renderAllTables();
  syncStateToUrl();
  ClipboardUtil.showToast(`+ ${type === 'pivot' ? 'Tabla Dinámica' : 'Tabla Plana'} agregada`, 'success');
}

/**
 * Renderiza todas las tablas en el contenedor principal
 */
function renderAllTables() {
  const container = document.getElementById('tablesContainer');
  const countBadge = document.getElementById('tableCountBadge');

  countBadge.innerText = `${AppState.tables.length} tabla${AppState.tables.length !== 1 ? 's' : ''}`;

  if (AppState.tables.length === 0) {
    container.innerHTML = `
      <div id="emptyWorkspaceState" class="empty-state">
        <i class="fa-solid fa-table-cells-large empty-icon"></i>
        <h3>No hay tablas creadas aún</h3>
        <p>Conecta un Google Sheet o carga los datos de demostración, luego haz clic en <strong>+ Tabla Dinámica</strong> o <strong>+ Tabla Plana</strong> para comenzar a analizar tu información.</p>
      </div>`;
    return;
  }

  container.innerHTML = '';

  AppState.tables.forEach((tableConfig, index) => {
    const tableCard = createTableCardElement(tableConfig, index);
    container.appendChild(tableCard);
  });
}

/**
 * Crea el elemento DOM para una tarjeta de tabla con sus controles y visor
 */
function createTableCardElement(config, index) {
  const card = document.createElement('div');
  card.className = 'table-card';
  card.id = `card_${config.id}`;

  const headers = AppState.dataset.headers || [];

  // 1. Header de la Tarjeta
  const headerHtml = `
    <div class="table-card-header">
      <div class="table-title-area">
        <span class="badge ${config.type === 'pivot' ? 'badge-success' : 'badge-info'}">
          <i class="fa-solid ${config.type === 'pivot' ? 'fa-table-cells' : 'fa-list-check'}"></i>
          ${config.type === 'pivot' ? 'Dinámica (Pivot)' : 'Plana / Filtros'}
        </span>
        <input 
          type="text" 
          class="table-title-input" 
          value="${escapeHtml(config.title)}" 
          placeholder="Nombre de la tabla..."
          onchange="updateTableField('${config.id}', 'title', this.value)"
        >
      </div>
      <div class="table-actions">
        <button class="btn btn-outline-primary btn-sm" onclick="copyEntireTable('${config.id}')" title="Copiar toda la tabla con formato Excel/Sheets">
          <i class="fa-solid fa-copy"></i> Copiar Tabla
        </button>
        <button class="btn btn-icon btn-sm" onclick="toggleCollapseTable('${config.id}')" title="Colapsar / Expandir">
          <i class="fa-solid ${config.isCollapsed ? 'fa-chevron-down' : 'fa-chevron-up'}"></i>
        </button>
        <button class="btn btn-icon btn-sm text-danger" onclick="deleteTable('${config.id}')" title="Eliminar tabla">
          <i class="fa-solid fa-trash"></i>
        </button>
      </div>
    </div>
  `;

  // 2. Panel de Configuración
  let configPanelHtml = `<div class="table-config-panel ${config.isCollapsed ? 'hidden' : ''}">`;

  if (config.type === 'pivot') {
    // Configuración para Tabla Dinámica
    configPanelHtml += `
      <div class="config-grid">
        <!-- Agrupar por Filas -->
        <div class="config-field">
          <label><i class="fa-solid fa-arrows-up-down"></i> Filas (Agrupar):</label>
          <div class="select-multi-box">
            ${headers.map(h => `
              <label class="checkbox-item">
                <input type="checkbox" value="${h}" ${config.rowFields.includes(h) ? 'checked' : ''} onchange="handlePivotDimensionChange('${config.id}', 'rowFields', '${h}', this.checked)">
                ${h}
              </label>
            `).join('')}
          </div>
        </div>

        <!-- Agrupar por Columnas -->
        <div class="config-field">
          <label><i class="fa-solid fa-arrows-left-right"></i> Columnas (Pivot):</label>
          <div class="select-multi-box">
            ${headers.map(h => `
              <label class="checkbox-item">
                <input type="checkbox" value="${h}" ${config.colFields.includes(h) ? 'checked' : ''} onchange="handlePivotDimensionChange('${config.id}', 'colFields', '${h}', this.checked)">
                ${h}
              </label>
            `).join('')}
          </div>
        </div>

        <!-- Campo de Valor / Métrica -->
        <div class="config-field">
          <label><i class="fa-solid fa-calculator"></i> Campo de Valor:</label>
          <select class="form-control" onchange="updateTableField('${config.id}', 'valField', this.value)">
            ${headers.map(h => `<option value="${h}" ${config.valField === h ? 'selected' : ''}>${h}</option>`).join('')}
          </select>
        </div>

        <!-- Función de Agregación -->
        <div class="config-field">
          <label><i class="fa-solid fa-function"></i> Agregación:</label>
          <select class="form-control" onchange="updateTableField('${config.id}', 'aggFunc', this.value)">
            <option value="SUM" ${config.aggFunc === 'SUM' ? 'selected' : ''}>Suma (SUM)</option>
            <option value="COUNT" ${config.aggFunc === 'COUNT' ? 'selected' : ''}>Conteo (COUNT)</option>
            <option value="AVG" ${config.aggFunc === 'AVG' ? 'selected' : ''}>Promedio (AVG)</option>
            <option value="MIN" ${config.aggFunc === 'MIN' ? 'selected' : ''}>Mínimo (MIN)</option>
            <option value="MAX" ${config.aggFunc === 'MAX' ? 'selected' : ''}>Máximo (MAX)</option>
          </select>
        </div>

        <!-- Buscador Global de Tabla -->
        <div class="config-field">
          <label><i class="fa-solid fa-magnifying-glass"></i> Buscar en tabla:</label>
          <input 
            type="text" 
            class="form-control" 
            placeholder="Filtrar texto..." 
            value="${escapeHtml(config.globalSearch || '')}"
            oninput="updateTableField('${config.id}', 'globalSearch', this.value)"
          >
        </div>
      </div>
    `;
  } else {
    // Configuración para Tabla Plana
    configPanelHtml += `
      <div class="config-grid">
        <!-- Selector de Columnas Visibles -->
        <div class="config-field" style="grid-column: span 2;">
          <label><i class="fa-solid fa-table-columns"></i> Columnas Visibles:</label>
          <div class="select-multi-box" style="display: flex; flex-direction: row; flex-wrap: wrap; gap: 0.75rem;">
            ${headers.map(h => `
              <label class="checkbox-item">
                <input type="checkbox" value="${h}" ${(config.selectedColumns || []).includes(h) ? 'checked' : ''} onchange="handleFlatColumnToggle('${config.id}', '${h}', this.checked)">
                ${h}
              </label>
            `).join('')}
          </div>
        </div>

        <!-- Ordenar por -->
        <div class="config-field">
          <label><i class="fa-solid fa-arrow-down-a-z"></i> Ordenar por:</label>
          <div style="display: flex; gap: 0.35rem;">
            <select class="form-control" onchange="updateTableField('${config.id}', 'sortBy', this.value)">
              <option value="">(Sin orden)</option>
              ${headers.map(h => `<option value="${h}" ${config.sortBy === h ? 'selected' : ''}>${h}</option>`).join('')}
            </select>
            <select class="form-control" style="width: 80px;" onchange="updateTableField('${config.id}', 'sortOrder', this.value)">
              <option value="asc" ${config.sortOrder === 'asc' ? 'selected' : ''}>ASC</option>
              <option value="desc" ${config.sortOrder === 'desc' ? 'selected' : ''}>DESC</option>
            </select>
          </div>
        </div>

        <!-- Buscador Global -->
        <div class="config-field">
          <label><i class="fa-solid fa-magnifying-glass"></i> Buscar en tabla:</label>
          <input 
            type="text" 
            class="form-control" 
            placeholder="Filtrar texto..." 
            value="${escapeHtml(config.globalSearch || '')}"
            oninput="updateTableField('${config.id}', 'globalSearch', this.value)"
          >
        </div>
      </div>
    `;
  }

  // Sección de Filtros Específicos por Fila/Columna
  configPanelHtml += `
    <div class="filters-container" style="margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px dashed var(--border-color);">
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.4rem;">
        <span style="font-size: 0.8rem; font-weight: 600; color: var(--text-secondary);">
          <i class="fa-solid fa-filter"></i> Filtros Específicos (${(config.filters || []).length}):
        </span>
        <button class="btn btn-xs btn-outline-primary" onclick="addFilterRow('${config.id}')">
          <i class="fa-solid fa-plus"></i> Añadir Filtro
        </button>
      </div>

      <div class="filters-list" id="filters_list_${config.id}">
        ${(config.filters || []).map((f, fIdx) => `
          <div class="filter-row">
            <select class="form-control" onchange="updateFilterField('${config.id}', ${fIdx}, 'field', this.value)">
              ${headers.map(h => `<option value="${h}" ${f.field === h ? 'selected' : ''}>${h}</option>`).join('')}
            </select>
            <select class="form-control" onchange="updateFilterField('${config.id}', ${fIdx}, 'operator', this.value)">
              <option value="contains" ${f.operator === 'contains' ? 'selected' : ''}>contiene</option>
              <option value="equals" ${f.operator === 'equals' ? 'selected' : ''}>es igual a</option>
              <option value="starts_with" ${f.operator === 'starts_with' ? 'selected' : ''}>empieza con</option>
              <option value="not_equals" ${f.operator === 'not_equals' ? 'selected' : ''}>no es igual</option>
              <option value="gt" ${f.operator === 'gt' ? 'selected' : ''}>mayor que (&gt;)</option>
              <option value="gte" ${f.operator === 'gte' ? 'selected' : ''}>mayor o igual (&ge;)</option>
              <option value="lt" ${f.operator === 'lt' ? 'selected' : ''}>menor que (&lt;)</option>
              <option value="lte" ${f.operator === 'lte' ? 'selected' : ''}>menor o igual (&le;)</option>
            </select>
            <input 
              type="text" 
              class="form-control" 
              placeholder="Valor..." 
              value="${escapeHtml(f.value || '')}"
              oninput="updateFilterField('${config.id}', ${fIdx}, 'value', this.value)"
            >
            <button class="btn btn-xs btn-outline-danger" onclick="removeFilterRow('${config.id}', ${fIdx})" title="Quitar filtro">
              <i class="fa-solid fa-xmark"></i>
            </button>
          </div>
        `).join('')}
      </div>
    </div>
  `;

  configPanelHtml += `</div>`;

  // 3. Renderizado de Datos de la Tabla
  let tableDataHtml = '';
  let footerInfo = '';

  if (config.type === 'pivot') {
    const pivotResult = PivotEngine.buildPivotData(AppState.dataset.rows, config);
    tableDataHtml = PivotEngine.renderTableHtml({
      id: config.id,
      type: 'pivot',
      headers: pivotResult.headers,
      matrixRows: pivotResult.matrixRows,
      grandTotalRow: pivotResult.grandTotalRow
    });
    footerInfo = `<span>Registros procesados: <strong>${pivotResult.rowCount}</strong></span><span>Haz clic en cualquier celda para copiar su valor</span>`;
  } else {
    const flatResult = PivotEngine.buildFlatData(AppState.dataset.rows, config);
    tableDataHtml = PivotEngine.renderTableHtml({
      id: config.id,
      type: 'flat',
      headers: flatResult.headers,
      rows: flatResult.rows
    });
    footerInfo = `<span>Mostrando <strong>${flatResult.rows.length}</strong> de <strong>${flatResult.totalCount}</strong> filas</span><span>Haz clic en un encabezado para copiar la columna</span>`;
  }

  // 4. Footer de la Tarjeta
  const footerHtml = `
    <div class="table-footer">
      ${footerInfo}
    </div>
  `;

  card.innerHTML = headerHtml + configPanelHtml + tableDataHtml + footerHtml;
  return card;
}

/**
 * Manejadores de cambios en la configuración de tablas
 */
window.updateTableField = function(tableId, field, value) {
  const table = AppState.tables.find(t => t.id === tableId);
  if (!table) return;

  table[field] = value;
  reRenderSingleTable(tableId);
  syncStateToUrl();
};

window.handlePivotDimensionChange = function(tableId, dimType, fieldName, isChecked) {
  const table = AppState.tables.find(t => t.id === tableId);
  if (!table) return;

  if (isChecked) {
    if (!table[dimType].includes(fieldName)) {
      table[dimType].push(fieldName);
    }
  } else {
    table[dimType] = table[dimType].filter(f => f !== fieldName);
  }

  reRenderSingleTable(tableId);
  syncStateToUrl();
};

window.handleFlatColumnToggle = function(tableId, fieldName, isChecked) {
  const table = AppState.tables.find(t => t.id === tableId);
  if (!table) return;

  if (isChecked) {
    if (!table.selectedColumns.includes(fieldName)) {
      table.selectedColumns.push(fieldName);
    }
  } else {
    table.selectedColumns = table.selectedColumns.filter(f => f !== fieldName);
  }

  reRenderSingleTable(tableId);
  syncStateToUrl();
};

window.addFilterRow = function(tableId) {
  const table = AppState.tables.find(t => t.id === tableId);
  if (!table) return;

  if (!table.filters) table.filters = [];
  const defaultHeader = AppState.dataset.headers[0] || '';
  table.filters.push({ field: defaultHeader, operator: 'contains', value: '' });

  reRenderSingleTable(tableId);
  syncStateToUrl();
};

window.updateFilterField = function(tableId, filterIndex, key, val) {
  const table = AppState.tables.find(t => t.id === tableId);
  if (!table || !table.filters || !table.filters[filterIndex]) return;

  table.filters[filterIndex][key] = val;
  reRenderSingleTable(tableId);
  syncStateToUrl();
};

window.removeFilterRow = function(tableId, filterIndex) {
  const table = AppState.tables.find(t => t.id === tableId);
  if (!table || !table.filters) return;

  table.filters.splice(filterIndex, 1);
  reRenderSingleTable(tableId);
  syncStateToUrl();
};

window.deleteTable = function(tableId) {
  AppState.tables = AppState.tables.filter(t => t.id !== tableId);
  renderAllTables();
  syncStateToUrl();
  ClipboardUtil.showToast('Tabla eliminada', 'info');
};

window.toggleCollapseTable = function(tableId) {
  const table = AppState.tables.find(t => t.id === tableId);
  if (!table) return;

  table.isCollapsed = !table.isCollapsed;
  reRenderSingleTable(tableId);
};

window.copyEntireTable = function(tableId) {
  const tableEl = document.getElementById(`table_${tableId}`);
  if (tableEl && window.ClipboardUtil) {
    window.ClipboardUtil.copyTable(tableEl);
  }
};

/**
 * Re-renderiza una única tarjeta de tabla sin recargar todo el DOM
 */
function reRenderSingleTable(tableId) {
  const card = document.getElementById(`card_${tableId}`);
  const config = AppState.tables.find(t => t.id === tableId);
  if (!card || !config) {
    renderAllTables();
    return;
  }

  const newCard = createTableCardElement(config, AppState.tables.indexOf(config));
  card.replaceWith(newCard);
}

/**
 * Sincroniza el estado actual en la URL comprimida con LZ-String
 */
function syncStateToUrl(notify = false) {
  const stateToPersist = {
    sheetId: AppState.sheetId,
    sheetUrl: AppState.sheetUrl,
    activeTab: AppState.activeTab,
    isDemo: AppState.isDemo,
    tables: AppState.tables
  };

  UrlStateManager.persistToUrl(stateToPersist);

  if (notify) {
    UrlStateManager.copyShareableUrl(stateToPersist);
  }
}

/**
 * Restaura el estado desde el hash de la URL si existe
 */
async function restoreStateFromUrl() {
  const savedState = UrlStateManager.decodeState();
  if (!savedState) return;

  if (savedState.sheetUrl) {
    document.getElementById('sheetUrlInput').value = savedState.sheetUrl;
    AppState.sheetUrl = savedState.sheetUrl;
    AppState.sheetId = savedState.sheetId;
  }

  if (savedState.activeTab) {
    AppState.activeTab = savedState.activeTab;
  }

  if (savedState.tables && Array.isArray(savedState.tables)) {
    AppState.tables = savedState.tables;
  }

  if (savedState.isDemo) {
    loadDemoDataset();
    return;
  }

  // Si había un sheet configurado y el usuario está autenticado, cargarlo automáticamente
  if (AppState.sheetId && GoogleSheetsService.isAuthenticated()) {
    await performFetchTabs(AppState.sheetId, AppState.sheetUrl);
  } else if (AppState.sheetId) {
    ClipboardUtil.showToast('Se restauró la configuración de tablas. Inicia sesión con Google para cargar los datos del Sheet.', 'info', 5000);
    renderAllTables();
  }
}

/**
 * Utilidad de escape para prevenir XSS en inputs
 */
function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
