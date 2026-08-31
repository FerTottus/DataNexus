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
      window.rawFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!A5:P250");
      window.rawSecos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!R5:AH250");

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

    // 1. Detectar dinámicamente cuál es la columna del Tiempo (Ej. "[6-2026]" o "32-2025")
    let timeColName = null;
    for (let h of rawData.headers) {
      // Tomamos una muestra de las primeras 5 filas a ver si alguna tiene un formato de fecha/semana
      for(let i=0; i<Math.min(5, rawData.rows.length); i++) {
        let val = String(rawData.rows[i]?.[h] || '');
        if (val.includes('-202') || (val.includes('[') && val.includes(']'))) {
          timeColName = h;
          break;
        }
      }
      if(timeColName) break;
    }
    // Fallbacks si la autodetección falla
    if (!timeColName) {
       timeColName = rawData.headers.find(h => h.toUpperCase() === 'SEMANA2' || h.toUpperCase() === 'SEMANA' || h.toUpperCase() === 'SEMAÑO');
    }
    if (!timeColName) timeColName = rawData.headers[1] || rawData.headers[0]; // Último recurso

    // 2. Extraer solo filas válidas
    let validRows = rawData.rows.filter(r => r[timeColName] && String(r[timeColName]).trim() !== '');

    // Para evitar filas de "Total general", filtramos filas donde la fecha diga Total
    validRows = validRows.filter(r => !String(r[timeColName]).toUpperCase().includes('TOTAL'));

    let rowsToPlot = validRows;

    // 3. Lógica de Filtro: Últimas X semanas del año actual + Últimas X semanas del año anterior
    if (filterValue !== 'all') {
        const numWeeks = parseInt(filterValue, 10);
        
        // Encontrar el año máximo en los datos
        let maxYear = 0;
        validRows.forEach(r => {
            const match = String(r[timeColName]).match(/-(\d{4})/);
            if (match) {
                const y = parseInt(match[1]);
                if (y > maxYear) maxYear = y;
            }
        });
        
        if (maxYear > 0) {
            // Obtener filas del año actual y tomar las últimas X
            const currentYearRows = validRows.filter(r => String(r[timeColName]).includes(`-${maxYear}`));
            const lastWeeksCurrentYear = currentYearRows.slice(-numWeeks);
            
            // Buscar su equivalente en el año anterior
            const previousYearRows = [];
            lastWeeksCurrentYear.forEach(cr => {
                const prevLabel = String(cr[timeColName]).replace(`-${maxYear}`, `-${maxYear - 1}`);
                const pr = validRows.find(r => String(r[timeColName]) === prevLabel);
                if (pr) {
                    previousYearRows.push(pr);
                }
            });
            
            // Unir secuencialmente
            rowsToPlot = [...previousYearRows, ...lastWeeksCurrentYear];
        } else {
            // Si no detectó año, simplemente tomar las ultimas N * 2 filas por si acaso
            rowsToPlot = validRows.slice(-(numWeeks * 2));
        }
    }

    // Si luego del filtro no hay datos, dibujar vacío
    if(rowsToPlot.length === 0) rowsToPlot = validRows;

    // 4. Mapeo Flexible de Columnas (busca nombres que contengan las palabras clave)
    const getCol = (keywords) => rawData.headers.find(h => keywords.some(k => h.toUpperCase().includes(k)));

    const colInventario = getCol(['INVENTARIO ACT', 'INVENTARIO']);
    const colRecibo = getCol(['RECIBO', 'INGRESO']);
    const colDespacho = getCol(['DESPACHO', 'SALIDA']);
    
    // Si la data no tiene "PLAN INV", usamos el nombre tal cual
    const colPlanInv = getCol(['PLAN INV', 'PLAN INVENTARIO']);
    const colPlanRecibo = getCol(['PLAN RECIBO', 'PLAN INGRESO']);
    const colPlanDespacho = getCol(['PLAN DESPACHO', 'PLAN SALIDA']);

    const labels = rowsToPlot.map(r => r[timeColName]);
    
    const inventario = rowsToPlot.map(r => parseFloat(r[colInventario]) || 0);
    const recibo = rowsToPlot.map(r => parseFloat(r[colRecibo]) || 0);
    const despacho = rowsToPlot.map(r => parseFloat(r[colDespacho]) || 0);
    const planInv = colPlanInv ? rowsToPlot.map(r => parseFloat(r[colPlanInv]) || 0) : [];
    const planRecibo = colPlanRecibo ? rowsToPlot.map(r => parseFloat(r[colPlanRecibo]) || 0) : [];
    const planDespacho = colPlanDespacho ? rowsToPlot.map(r => parseFloat(r[colPlanDespacho]) || 0) : [];

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
        type: 'bar',
        label: 'RECIBO',
        data: data.recibo,
        backgroundColor: '#457b9d', // Azul Barra
        borderColor: '#1d3557',
        borderWidth: 1,
        order: 5
      },
      {
        type: 'bar',
        label: 'DESPACHO',
        data: data.despacho,
        backgroundColor: '#e76f51', // Naranja Barra
        borderColor: '#d00000',
        borderWidth: 1,
        order: 6
      }
    ];

    if(data.planInv && data.planInv.length > 0 && data.planInv.some(v => v > 0)) {
        datasets.push({
            type: 'line',
            label: 'PLAN INV',
            data: data.planInv,
            borderColor: '#2a9d8f', // Verde Plan
            borderWidth: 2,
            borderDash: [5, 5],
            fill: false,
            pointRadius: 0,
            order: 2
        });
    }

    if(data.planRecibo && data.planRecibo.length > 0 && data.planRecibo.some(v => v > 0)) {
        datasets.push({
            type: 'line',
            label: 'PLAN RECIBO',
            data: data.planRecibo,
            borderColor: '#1d3557', // Azul Oscuro Plan
            borderWidth: 2,
            borderDash: [5, 5],
            fill: false,
            pointRadius: 0,
            order: 3
        });
    }

    if(data.planDespacho && data.planDespacho.length > 0 && data.planDespacho.some(v => v > 0)) {
        datasets.push({
            type: 'line',
            label: 'PLAN DESPACHO',
            data: data.planDespacho,
            borderColor: '#9d0208', // Rojo Oscuro Plan
            borderWidth: 2,
            borderDash: [5, 5],
            fill: false,
            pointRadius: 0,
            order: 4
        });
    }

    charts[canvasId] = new Chart(ctx, {
      data: {
        labels: labels,
        datasets: datasets
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
    const validHeaders = headers.filter(h => h && h.trim() !== '');

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
