document.addEventListener('DOMContentLoaded', () => {
  if (window.GoogleSheetsService) {
    window.GoogleSheetsService.initAuth();
    updateAuthUI();
  }

  const btnGoogleLogin = document.getElementById('btnGoogleLogin');
  const btnGoogleLogout = document.getElementById('btnGoogleLogout');
  const btnFetchData = document.getElementById('btnFetchData');
  const btnBrowseDrive = document.getElementById('btnBrowseDrive');
  const sheetUrlInput = document.getElementById('sheetUrlInput');
  const loadingIndicator = document.getElementById('loadingIndicator');
  const dashboardSection = document.getElementById('dashboardSection');
  const connectBox = document.getElementById('connectBox');
  const connectionSuccessInfo = document.getElementById('connectionSuccessInfo');
  const btnChangeSheet = document.getElementById('btnChangeSheet');
  const tabBtns = document.querySelectorAll('.tab-btn');
  const btnApplyFilters = document.getElementById('btnApplyFilters');

  btnGoogleLogin?.addEventListener('click', () => {
    window.GoogleSheetsService.requestAccessToken((success) => {
      if (success) updateAuthUI();
    });
  });

  btnGoogleLogout?.addEventListener('click', () => {
    window.GoogleSheetsService.logout();
    updateAuthUI();
  });

  function updateAuthUI() {
    const isAuth = window.GoogleSheetsService.isAuthenticated();
    document.getElementById('btnGoogleLogin')?.classList.toggle('hidden', isAuth);
    document.getElementById('userInfo')?.classList.toggle('hidden', !isAuth);
    if (isAuth && window.GoogleSheetsService.userEmail) {
      document.getElementById('userEmail').textContent = window.GoogleSheetsService.userEmail;
    }
  }

  // Explorador Drive
  btnBrowseDrive?.addEventListener('click', () => {
    if (!window.GoogleSheetsService.isAuthenticated()) {
      window.GoogleSheetsService.requestAccessToken((success) => {
        if (success) { updateAuthUI(); openDriveModal(); }
      });
    } else {
      openDriveModal();
    }
  });

  document.getElementById('close-modal-btn')?.addEventListener('click', () => {
    document.getElementById('drive-modal').classList.add('hidden');
  });

  document.getElementById('drive-modal')?.addEventListener('click', (e) => {
    if (e.target.id === 'drive-modal') e.target.classList.add('hidden');
  });

  async function openDriveModal() {
    document.getElementById('drive-modal').classList.remove('hidden');
    const fileList = document.getElementById('drive-file-list');
    const loading = document.getElementById('drive-loading');
    const emptyMsg = document.getElementById('drive-empty');
    
    fileList.innerHTML = '';
    loading.classList.remove('hidden');
    emptyMsg.classList.add('hidden');

    const files = await window.GoogleSheetsService.fetchRecentSpreadsheets();
    loading.classList.add('hidden');

    if (!files || files.length === 0) {
      emptyMsg.classList.remove('hidden');
      return;
    }
    window._driveFiles = files;
    renderDriveFiles(files);
  }

  function renderDriveFiles(files) {
    const fileList = document.getElementById('drive-file-list');
    fileList.innerHTML = '';
    files.forEach(file => {
      const li = document.createElement('li');
      li.className = 'file-item';
      li.style.cursor = 'pointer';
      li.style.padding = '10px';
      li.style.borderBottom = '1px solid #eee';
      li.innerHTML = `<i class="fa-solid fa-file-excel text-success"></i> <span style="margin-left:8px;">${file.name}</span>`;
      li.addEventListener('click', () => {
        sheetUrlInput.value = file.id;
        document.getElementById('drive-modal').classList.add('hidden');
        btnFetchData.click();
      });
      fileList.appendChild(li);
    });
  }

  window.filterDriveList = function(text) {
    if (!window._driveFiles) return;
    const lower = text.toLowerCase();
    const filtered = window._driveFiles.filter(f => f.name.toLowerCase().includes(lower));
    renderDriveFiles(filtered);
  };

  // Tabs
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.target).classList.add('active');
    });
  });

  btnChangeSheet?.addEventListener('click', () => {
    connectBox.classList.remove('hidden');
    connectionSuccessInfo.classList.add('hidden');
  });

  // Fetch Data
  btnFetchData?.addEventListener('click', async () => {
    const urlOrId = sheetUrlInput.value.trim();
    if (!urlOrId) return alert("Por favor, ingresa la URL o busca un archivo.");
    const sheetId = window.GoogleSheetsService.extractSpreadsheetId(urlOrId);
    if (!sheetId) return alert("ID inválido.");
    if (!window.GoogleSheetsService.isAuthenticated()) return alert("Inicia sesión primero.");

    loadingIndicator.classList.remove('hidden');
    
    try {
      // BD_Grafico contiene los totales por fecha
      window.rawFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!A5:P150");
      window.rawSecos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!R5:AH150");

      connectBox.classList.add('hidden');
      connectionSuccessInfo.classList.remove('hidden');
      document.getElementById('connectedSheetName').textContent = `Documento cargado (ID: ${sheetId.substring(0, 8)}...)`;
      dashboardSection.classList.remove('hidden');

      applyFiltersAndRender();
    } catch (e) {
      console.error(e);
      alert("Error: " + e.message);
    } finally {
      loadingIndicator.classList.add('hidden');
    }
  });

  btnApplyFilters?.addEventListener('click', () => {
    applyFiltersAndRender();
  });

  function applyFiltersAndRender() {
    if(!window.rawFrescos || !window.rawSecos) return;
    const filterValue = document.getElementById('weeksFilter').value;
    
    processAndRenderSection('Secos', window.rawSecos, filterValue);
    processAndRenderSection('Frescos', window.rawFrescos, filterValue);
  }

  // --- LÓGICA DE PROCESAMIENTO Y GRÁFICOS ---
  const charts = {};

  function processAndRenderSection(prefix, rawData, filterValue) {
    if(!rawData.rows || rawData.rows.length === 0) return;

    // Las filas tienen el SemAño, DESPACHO, RECIBO, etc.
    // Ej: row['SemAño'] = '[6-2026]'
    
    // Filtrar solo filas con SemAño válido
    let validRows = rawData.rows.filter(r => r['SemAño'] && String(r['SemAño']).includes('-'));
    let rowsToPlot = validRows;

    // Lógica de Filtro: Últimas X semanas del año actual + Últimas X semanas del año anterior (secuencial)
    if (filterValue !== 'all') {
        const numWeeks = parseInt(filterValue, 10);
        
        let maxYear = 0;
        validRows.forEach(r => {
            const match = String(r['SemAño']).match(/-(\d{4})/);
            if (match) {
                const y = parseInt(match[1]);
                if (y > maxYear) maxYear = y;
            }
        });
        
        // Obtener filas del año actual y tomar las últimas X
        const currentYearRows = validRows.filter(r => String(r['SemAño']).includes(`-${maxYear}`));
        const lastWeeksCurrentYear = currentYearRows.slice(-numWeeks);
        
        // Buscar su equivalente en el año anterior
        const previousYearRows = [];
        lastWeeksCurrentYear.forEach(cr => {
            const prevLabel = String(cr['SemAño']).replace(`-${maxYear}`, `-${maxYear - 1}`);
            const pr = validRows.find(r => String(r['SemAño']) === prevLabel);
            if (pr) {
                previousYearRows.push(pr);
            }
        });
        
        // Unir secuencialmente: primero las del año anterior, luego las de este año
        rowsToPlot = [...previousYearRows, ...lastWeeksCurrentYear];
    }

    // Extraer Vectores de Datos para el Gráfico
    const labels = rowsToPlot.map(r => r['SemAño']);
    const inventario = rowsToPlot.map(r => parseFloat(r['INVENTARIO ACT']) || parseFloat(r['INVENTARIO']) || 0);
    const recibo = rowsToPlot.map(r => parseFloat(r['RECIBO']) || 0);
    const despacho = rowsToPlot.map(r => parseFloat(r['DESPACHO']) || 0);
    const planInv = rowsToPlot.map(r => parseFloat(r['PLAN INV']) || 0);
    const planRecibo = rowsToPlot.map(r => parseFloat(r['PLAN RECIBO']) || 0);
    const planDespacho = rowsToPlot.map(r => parseFloat(r['PLAN DESPACHO']) || 0);

    renderMixedChart(prefix, labels, { inventario, recibo, despacho, planInv, planRecibo, planDespacho });
    renderDataTable(prefix, rawData.headers, rowsToPlot);
  }

  function renderMixedChart(prefix, labels, data) {
    const canvasId = `chart${prefix}`;
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    const datasets = [
      {
        type: 'line',
        label: 'INVENTARIO',
        data: data.inventario,
        backgroundColor: 'rgba(169, 169, 169, 0.4)', // Gris Área
        borderColor: '#7a7a7a',
        borderWidth: 2,
        fill: true,
        order: 1
      },
      {
        type: 'line',
        label: 'PLAN INV',
        data: data.planInv,
        borderColor: '#2a9d8f', // Verde Plan
        borderWidth: 2,
        borderDash: [5, 5],
        fill: false,
        pointRadius: 0,
        order: 2
      },
      {
        type: 'line',
        label: 'PLAN RECIBO',
        data: data.planRecibo,
        borderColor: '#457b9d', // Azul Plan
        borderWidth: 2,
        borderDash: [5, 5],
        fill: false,
        pointRadius: 0,
        order: 3
      },
      {
        type: 'line',
        label: 'PLAN DESPACHO',
        data: data.planDespacho,
        borderColor: '#e76f51', // Naranja Plan
        borderWidth: 2,
        borderDash: [5, 5],
        fill: false,
        pointRadius: 0,
        order: 4
      },
      {
        type: 'bar',
        label: 'RECIBO',
        data: data.recibo,
        backgroundColor: '#6baed6', // Azul Barra
        borderColor: '#3182bd',
        borderWidth: 1,
        order: 5
      },
      {
        type: 'bar',
        label: 'DESPACHO',
        data: data.despacho,
        backgroundColor: '#fd8d3c', // Naranja Barra
        borderColor: '#e6550d',
        borderWidth: 1,
        order: 6
      }
    ];

    charts[canvasId] = new Chart(ctx, {
      data: {
        labels: labels,
        datasets: datasets.filter(ds => ds.data.some(val => val > 0)) // Solo graficar los que tengan datos
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { 
            position: 'top',
            labels: { usePointStyle: true, boxWidth: 10 }
          },
          tooltip: { mode: 'index', intersect: false }
        },
        scales: {
          x: {
             stacked: false,
             grid: { display: false }
          },
          y: {
            stacked: false,
            beginAtZero: true,
            ticks: {
               callback: function(value) {
                  return value.toLocaleString('es-PE');
               }
            }
          }
        },
        interaction: { mode: 'index', intersect: false }
      }
    });
  }

  function renderDataTable(prefix, headers, rows) {
    const thead = document.getElementById(`table${prefix}Header`);
    const tbody = document.getElementById(`table${prefix}Body`);
    thead.innerHTML = '';
    tbody.innerHTML = '';

    // Filtrar columnas vacías
    const validHeaders = headers.filter(h => h.trim() !== '');

    validHeaders.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        thead.appendChild(th);
    });

    rows.forEach(row => {
        const tr = document.createElement('tr');
        validHeaders.forEach(h => {
            const td = document.createElement('td');
            const val = row[h];
            if(typeof val === 'number' || (typeof val === 'string' && !isNaN(parseFloat(val)) && val.trim() !== '')) {
                td.textContent = parseFloat(val).toLocaleString('es-PE', { maximumFractionDigits: 0 });
            } else {
                td.textContent = val || '';
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
  }
});
