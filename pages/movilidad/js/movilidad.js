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
}

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
    // Descargar las 3 hojas concurrentemente
    const [secos, ppa, frescos] = await Promise.all([
      GoogleSheetsService.fetchSheetData(sheetId, 'BD SECOS').catch(() => null),
      GoogleSheetsService.fetchSheetData(sheetId, 'BD PPA').catch(() => null),
      GoogleSheetsService.fetchSheetData(sheetId, 'BD FRESCOS').catch(() => null)
    ]);

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
          ruta: String(row['RUTA PARADERO'] || '').trim()
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

    document.getElementById('dashboardFilters').classList.remove('hidden');
    document.getElementById('dashboardContent').classList.remove('hidden');

    applyFilters();

    if(window.ClipboardUtil) ClipboardUtil.showToast('Datos cargados exitosamente', 'success');

  } catch (err) {
    console.error(err);
    badge.className = 'badge badge-danger';
    badge.innerText = 'Error de conexión';
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
