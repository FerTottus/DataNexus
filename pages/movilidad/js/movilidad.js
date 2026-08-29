/**
 * Script Principal para el Dashboard de Movilidad
 */

// Estado global de la aplicación
const AppState = {
  sheetId: '',
  sheetUrl: '',
  rawEmployees: [], // Todos los empleados combinados
  filteredEmployees: [], // Empleados después de aplicar filtros
  charts: {} // Referencias a instancias de Chart.js
};

document.addEventListener('DOMContentLoaded', () => {
  initUIEvents();
  GoogleSheetsService.initAuth(); // Restaura token
  checkAuthAndConfig();
  initCharts();
});

function initUIEvents() {
  document.getElementById('btnGoogleLogin').addEventListener('click', () => {
    GoogleSheetsService.requestAccessToken((success) => {
      if (success) updateAuthUI(true);
    });
  });

  document.getElementById('btnGoogleLogout').addEventListener('click', () => {
    GoogleSheetsService.logout();
    updateAuthUI(false);
  });

  document.getElementById('btnFetchData').addEventListener('click', () => {
    fetchData();
  });

  document.getElementById('btnApplyFilters').addEventListener('click', () => {
    applyFilters();
  });

  // Explorar Drive
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

  // Modal close
  document.getElementById('close-modal-btn').addEventListener('click', () => {
    document.getElementById('drive-modal').classList.add('hidden');
  });

  document.getElementById('drive-modal').addEventListener('click', (e) => {
    if (e.target.id === 'drive-modal') {
      e.target.classList.add('hidden');
    }
  });

  // Toggle Connection Panel
  const toggleConnectionBtn = document.getElementById('toggleConnectionBtn');
  if (toggleConnectionBtn) {
    toggleConnectionBtn.addEventListener('click', () => {
      const body = document.getElementById('connectionCardBody');
      const icon = toggleConnectionBtn.querySelector('i');
      if (body.style.display === 'none') {
        body.style.display = 'block';
        icon.classList.replace('fa-chevron-down', 'fa-chevron-up');
      } else {
        body.style.display = 'none';
        icon.classList.replace('fa-chevron-up', 'fa-chevron-down');
      }
    });
  }

  // Toggle Filters Panel
  const toggleFiltersBtn = document.getElementById('toggleFiltersBtn');
  if (toggleFiltersBtn) {
    toggleFiltersBtn.addEventListener('click', () => {
      const panel = document.getElementById('dashboardFilters');
      const icon = toggleFiltersBtn.querySelector('i');
      if (panel.style.display === 'none') {
        panel.style.display = 'flex';
        icon.classList.replace('fa-chevron-down', 'fa-chevron-up');
      } else {
        panel.style.display = 'none';
        icon.classList.replace('fa-chevron-up', 'fa-chevron-down');
      }
    });
  }
}

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
    
    const date = new Date(file.modifiedTime).toLocaleDateString('es-ES', { day: 'numeric', month: 'short', year: 'numeric' });
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
    
    const excelWarningHtml = isExcel ? `<span class="file-badge" style="background: #fee2e2; color: #ef4444; border: 1px solid #fecaca; margin-right: 5px;" title="Ábrelo en Drive y dale a 'Guardar como Hoja de cálculo de Google'">⚠️ Inválido (.xlsx)</span>` : '';

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
      fetchData(); // Cargar automáticamente
    });
    
    list.appendChild(li);
  });
}

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

function checkAuthAndConfig() {
  const isAuth = GoogleSheetsService.isAuthenticated();
  updateAuthUI(isAuth);
}

function updateAuthUI(isAuthenticated) {
  const btnLogin = document.getElementById('btnGoogleLogin');
  const userInfo = document.getElementById('userInfo');
  const userEmailSpan = document.getElementById('userEmail');

  if (isAuthenticated) {
    btnLogin.classList.add('hidden');
    userInfo.classList.remove('hidden');
    userEmailSpan.innerText = GoogleSheetsService.userEmail || 'Conectado';
  } else {
    btnLogin.classList.remove('hidden');
    userInfo.classList.add('hidden');
  }
}

async function fetchData() {
  const input = document.getElementById('sheetUrlInput').value.trim();
  if (!input) {
    if(window.ClipboardUtil) ClipboardUtil.showToast('Ingresa la URL del Google Sheet', 'info');
    else alert('Ingresa la URL del Google Sheet');
    return;
  }

  const sheetId = GoogleSheetsService.extractSpreadsheetId(input);
  if (!sheetId) return;

  if (!GoogleSheetsService.isAuthenticated()) {
    GoogleSheetsService.requestAccessToken(async (success) => {
      if (success) {
        updateAuthUI(true);
        await loadAllSheets(sheetId);
      }
    });
    return;
  }

  await loadAllSheets(sheetId);
}

async function loadAllSheets(sheetId) {
  const badge = document.getElementById('connectionStatus');
  badge.className = 'badge badge-warning';
  badge.innerText = 'Descargando datos...';

  try {
    const [secos, ppa, frescos] = await Promise.all([
      GoogleSheetsService.fetchSheetData(sheetId, 'BD SECOS').catch(e => null),
      GoogleSheetsService.fetchSheetData(sheetId, 'BD PPA').catch(e => null),
      GoogleSheetsService.fetchSheetData(sheetId, 'BD FRESCOS').catch(e => null)
    ]);

    if (!secos && !ppa && !frescos) {
      throw new Error("No se pudo leer la información. Si tu archivo es un Excel (.xlsx), debes abrirlo en Google Drive y darle a 'Archivo > Guardar como Hoja de cálculo de Google'. También verifica que existan las pestañas 'BD SECOS', 'BD PPA' o 'BD FRESCOS'.");
    }

    let combined = [];

    const processSheet = (sheetData, areaName) => {
      if (!sheetData) return;
      sheetData.rows.forEach(row => {
        // Ignorar filas vacías o rotulos de información sin DNI o Nombres reales
        if(!row['Distrito']) return;
        
        combined.push({
          area: areaName,
          tipo: String(row['CLASIFICACION'] || '').toUpperCase().trim(),
          distrito: String(row['Distrito'] || '').trim(),
          distCd: parseFloat(row['DIST. AL CD (km)']) || 0,
          clasifCd: String(row['CLASIF. DIST. CD'] || '').trim(),
          distParadero: parseFloat(row['DIST. PARADERO (km)']) || 0,
          clasifParadero: String(row['CLASIF. DIST. PARADERO'] || '').trim(),
          ruta: String(row['RUTA PARADERO'] || '').trim(),
          paradero: String(row['PARADERO'] || row['PARADERO MÁS CERCANO'] || row['NOMBRE PARADERO'] || '').trim()
        });
      });
    };

    processSheet(secos, 'SECOS');
    processSheet(ppa, 'PPA');
    processSheet(frescos, 'FRESCOS');

    AppState.rawEmployees = combined;
    AppState.sheetId = sheetId;

    badge.className = 'badge badge-success';
    badge.innerText = 'Conectado - ' + combined.length + ' regs';

    document.getElementById('dashboardFiltersWrapper').classList.remove('hidden');
    document.getElementById('dashboardContent').classList.remove('hidden');

    applyFilters();

    if(window.ClipboardUtil) ClipboardUtil.showToast('Datos cargados exitosamente', 'success');

  } catch (err) {
    console.error(err);
    badge.className = 'badge badge-danger';
    badge.innerText = 'Error de conexión';
    if(window.ClipboardUtil) {
      ClipboardUtil.showToast(err.message, 'danger', 10000);
    } else {
      alert(err.message);
    }
  }
}

function applyFilters() {
  const filterArea = document.getElementById('filterArea').value;
  const filterTipo = document.getElementById('filterTipo').value;

  AppState.filteredEmployees = AppState.rawEmployees.filter(emp => {
    let matchArea = filterArea === 'TODOS' || emp.area === filterArea;
    let matchTipo = filterTipo === 'TODOS' || emp.tipo === filterTipo;
    return matchArea && matchTipo;
  });

  updateKPIs();
  updateCharts();
}

function updateKPIs() {
  const data = AppState.filteredEmployees;
  
  // Total Empleados
  document.getElementById('kpiTotal').innerText = data.length;

  // Promedios
  let sumDistCd = 0;
  let sumDistParadero = 0;
  let countValidCd = 0;
  let countValidParadero = 0;
  let distritosSet = new Set();

  data.forEach(emp => {
    if (emp.distCd > 0) {
      sumDistCd += emp.distCd;
      countValidCd++;
    }
    if (emp.distParadero > 0) {
      sumDistParadero += emp.distParadero;
      countValidParadero++;
    }
    if (emp.distrito) {
      distritosSet.add(emp.distrito);
    }
  });

  document.getElementById('kpiDistCD').innerText = countValidCd > 0 ? (sumDistCd / countValidCd).toFixed(2) : '0';
  document.getElementById('kpiDistParadero').innerText = countValidParadero > 0 ? (sumDistParadero / countValidParadero).toFixed(2) : '0';
  document.getElementById('kpiDistritos').innerText = distritosSet.size;
}

function initCharts() {
  // Chart defaults
  Chart.defaults.font.family = "'Inter', sans-serif";
  Chart.defaults.color = '#475569';
}

function updateCharts() {
  const data = AppState.filteredEmployees;

  // 1. Clasificación Distancia CD
  const countCd = {
    'Muy Cerca': 0, 'Cerca': 0, 'Moderada': 0, 'Lejos': 0
  };
  
  // 2. Clasificación Distancia Paraderos
  const countParadero = {
    'Muy Cerca': 0, 'Cerca': 0, 'Moderada': 0, 'Lejos': 0
  };

  // 3. Distritos (Top 10)
  const distritosCount = {};

  data.forEach(emp => {
    // CD
    if (emp.clasifCd.includes('Muy Cerca')) countCd['Muy Cerca']++;
    else if (emp.clasifCd.includes('Cerca')) countCd['Cerca']++;
    else if (emp.clasifCd.includes('Moderada')) countCd['Moderada']++;
    else if (emp.clasifCd.includes('Lejos')) countCd['Lejos']++;

    // Paradero
    if (emp.clasifParadero.includes('Muy Cerca')) countParadero['Muy Cerca']++;
    else if (emp.clasifParadero.includes('Cerca')) countParadero['Cerca']++;
    else if (emp.clasifParadero.includes('Moderada')) countParadero['Moderada']++;
    else if (emp.clasifParadero.includes('Lejos')) countParadero['Lejos']++;

    // Distrito
    if (emp.distrito) {
      distritosCount[emp.distrito] = (distritosCount[emp.distrito] || 0) + 1;
    }
  });

  const distritosSorted = Object.keys(distritosCount)
    .map(k => ({ name: k, val: distritosCount[k] }))
    .sort((a, b) => b.val - a.val)
    .slice(0, 10);

  // Calcular Rutas Disponibles
  const rutasCount = {};
  const paraderosCount = {};
  
  AppState.filteredEmployees.forEach(emp => {
    if (emp.ruta) rutasCount[emp.ruta] = (rutasCount[emp.ruta] || 0) + 1;
    if (emp.paradero) paraderosCount[emp.paradero] = (paraderosCount[emp.paradero] || 0) + 1;
  });

  const totalEmps = AppState.filteredEmployees.length;

  const rutasSorted = Object.keys(rutasCount)
    .map(k => ({ name: k, val: rutasCount[k] }))
    .sort((a, b) => b.val - a.val);

  const paraderosSorted = Object.keys(paraderosCount)
    .map(k => ({ name: k, val: paraderosCount[k] }))
    .sort((a, b) => b.val - a.val)
    .slice(0, 10);

  // Colores corporativos (semáforo)
  const colors = {
    'Muy Cerca': '#22c55e', // Verde
    'Cerca': '#eab308',     // Amarillo
    'Moderada': '#f97316',  // Naranja
    'Lejos': '#ef4444'      // Rojo
  };

  renderPieChart('chartDistCD', countCd, colors);
  renderPieChart('chartDistParadero', countParadero, colors);
  renderBarChart('chartDistritos', distritosSorted);

  // Renderizar Tablas
  renderTable('tableDistritos', distritosSorted, totalEmps);
  renderTable('tableRutas', rutasSorted, totalEmps);
  renderTable('tableParaderos', paraderosSorted, totalEmps);
}

function renderTable(tableId, dataArr, total) {
  const tbody = document.querySelector(`#${tableId} tbody`);
  if (!tbody) return;
  tbody.innerHTML = '';

  if (dataArr.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted">No hay datos disponibles</td></tr>';
    return;
  }

  dataArr.forEach((item, index) => {
    const pct = total > 0 ? ((item.val / total) * 100).toFixed(2) + '%' : '0%';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${item.name}</td>
      <td>${item.val}</td>
      <td>${pct}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderPieChart(canvasId, dataMap, colors) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  
  const labels = Object.keys(dataMap);
  const data = labels.map(k => dataMap[k]);
  const bgColors = labels.map(k => colors[k] || '#cbd5e1');

  if (AppState.charts[canvasId]) {
    AppState.charts[canvasId].destroy();
  }

  AppState.charts[canvasId] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: bgColors,
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom' }
      }
    }
  });
}

function renderBarChart(canvasId, dataArr) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  
  const labels = dataArr.map(d => d.name);
  const data = dataArr.map(d => d.val);

  if (AppState.charts[canvasId]) {
    AppState.charts[canvasId].destroy();
  }

  AppState.charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Volumen de Empleados',
        data: data,
        backgroundColor: '#3b82f6',
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0 } }
      },
      plugins: {
        legend: { display: false }
      }
    }
  });
}
