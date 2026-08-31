document.addEventListener('DOMContentLoaded', () => {
  // Inicialización
  if (window.GoogleSheetsService) {
    window.GoogleSheetsService.initAuth();
    updateAuthUI();
  }

  // Elementos UI globales
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

  // Autenticación
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

  // Modal Explorador de Drive
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

  // Cargar Datos
  btnChangeSheet?.addEventListener('click', () => {
    connectBox.classList.remove('hidden');
    connectionSuccessInfo.classList.add('hidden');
  });

  btnFetchData?.addEventListener('click', async () => {
    const urlOrId = sheetUrlInput.value.trim();
    if (!urlOrId) return alert("Por favor, ingresa la URL del Google Sheet.");
    const sheetId = window.GoogleSheetsService.extractSpreadsheetId(urlOrId);
    if (!sheetId) return alert("ID inválido.");
    if (!window.GoogleSheetsService.isAuthenticated()) return alert("Inicia sesión primero.");

    loadingIndicator.classList.remove('hidden');
    
    try {
      window.rawFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!A5:L109");
      window.rawSecos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!R5:AA109");

      connectBox.classList.add('hidden');
      connectionSuccessInfo.classList.remove('hidden');
      document.getElementById('connectedSheetName').textContent = "ID: " + sheetId;
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
    const filterType = document.getElementById('timeFilter').value;
    
    processAndRenderSection('Frescos', window.rawFrescos, filterType);
    processAndRenderSection('Secos', window.rawSecos, filterType);
  }

  // --- LÓGICA DE PROCESAMIENTO Y GRÁFICOS ---
  const charts = {};

  function processAndRenderSection(prefix, rawData, filterType) {
    if(!rawData.rows || rawData.rows.length === 0) return;

    // Asumimos tablas apiladas: RECIBO, DESPACHO, INVENTARIO
    // Y asumimos X-Axis = Semanas/Meses (las columnas)
    
    const firstCol = rawData.headers[0];
    const timeCols = extractAndFilterTimeCols(rawData.headers, filterType);

    // Separar datos en 3 bloques
    const blocks = { RECIBO: [], DESPACHO: [], INVENTARIO: [] };
    let currentBlock = 'RECIBO'; // por defecto el primero

    rawData.rows.forEach(row => {
      const label = String(row[firstCol] || '').toUpperCase();
      if(label.includes('RECIBO') || label.includes('INGRESO')) currentBlock = 'RECIBO';
      else if(label.includes('DESPACHO') || label.includes('SALIDA')) currentBlock = 'DESPACHO';
      else if(label.includes('INVENTARIO')) currentBlock = 'INVENTARIO';
      else if(label && !label.includes('TOTAL') && !label.includes('CAJAS')) {
        blocks[currentBlock].push(row);
      }
    });

    renderDataChart(`${prefix}Recibo`, 'bar', timeCols, blocks.RECIBO, firstCol);
    renderDataChart(`${prefix}Despacho`, 'line', timeCols, blocks.DESPACHO, firstCol);
    renderDataChart(`${prefix}Inventario`, 'bar', timeCols, blocks.INVENTARIO, firstCol);

    renderFilteredTable(`table${prefix}`, rawData.headers, timeCols, blocks.RECIBO, blocks.DESPACHO, blocks.INVENTARIO, firstCol);
  }

  function extractAndFilterTimeCols(headers, filterType) {
    // Extraer columnas omitiendo la primera (etiquetas) y el total general
    let cols = headers.slice(1).filter(h => !h.toUpperCase().includes('TOTAL'));
    
    if (filterType === 'all') return cols;

    // Lógica simple: Últimos 2 meses => aprox últimas 8 columnas (si son semanas) o 2 (si son meses)
    // Como no sabemos si la tabla está agrupada por semanas o meses, tomamos las últimas 8 columnas
    // para representar "2 meses" (asumiendo formato semanal que es muy comun en throughput).
    // Si la data tiene el formato "[6-2025]" podemos parsear.
    
    if (filterType === 'last2months') {
        return cols.slice(-8); // Tomamos las ultimas 8 columnas (aprox 2 meses si es semanal)
    }

    if (filterType === 'last2months_yoy') {
        // Tratar de buscar el año anterior a las últimas columnas
        const recentCols = cols.slice(-8);
        const lastYearCols = [];
        
        recentCols.forEach(col => {
            // ej: "[6-2026]" -> buscamos "[6-2025]"
            let match = col.match(/(\d{4})/);
            if (match) {
                let year = parseInt(match[1]);
                let prevYearCol = col.replace(year.toString(), (year-1).toString());
                if(cols.includes(prevYearCol)) lastYearCols.push(prevYearCol);
            }
        });
        
        // Si falló el regex, regresamos simplemente las primeras 8 columnas asumiendo que son del año pasado
        return lastYearCols.length > 0 ? lastYearCols : cols.slice(0, 8); 
    }

    return cols;
  }

  function renderDataChart(canvasId, type, xLabels, rows, labelCol) {
    const ctx = document.getElementById('chart' + canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    const colors = [
      '#4285F4', '#34A853', '#FBBC05', '#EA4335', '#673AB7', '#3F51B5', '#009688', '#FF9800', '#795548', '#607D8B'
    ];

    // Datasets: Cada fila (Ej. J01, J02) es una línea/barra en el gráfico
    const datasets = rows.map((row, index) => {
      return {
        label: row[labelCol] || `División ${index}`,
        data: xLabels.map(col => parseFloat(row[col]) || 0),
        backgroundColor: colors[index % colors.length] + '80', // Opacidad 50%
        borderColor: colors[index % colors.length],
        borderWidth: 2,
        tension: 0.3,
        fill: type === 'line' ? false : true
      };
    });

    charts[canvasId] = new Chart(ctx, {
      type: type,
      data: {
        labels: xLabels, // Eje X = Tiempo
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right' }
        },
        scales: {
          x: { stacked: false },
          y: { beginAtZero: true }
        }
      }
    });
  }

  function renderFilteredTable(tableId, allHeaders, timeCols, rRows, dRows, iRows, labelCol) {
    const thead = document.getElementById(`${tableId}Header`);
    const tbody = document.getElementById(`${tableId}Body`);
    thead.innerHTML = '';
    tbody.innerHTML = '';

    // Headers: [Categoria] [Division] [Fechas...]
    const finalHeaders = ['CATEGORÍA', labelCol, ...timeCols];
    finalHeaders.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        thead.appendChild(th);
    });

    const appendRows = (rows, categoryName) => {
        rows.forEach(row => {
            const tr = document.createElement('tr');
            
            // Categoria
            const tdCat = document.createElement('td');
            tdCat.textContent = categoryName;
            tdCat.style.fontWeight = 'bold';
            tr.appendChild(tdCat);

            // Division
            const tdDiv = document.createElement('td');
            tdDiv.textContent = row[labelCol] || '';
            tr.appendChild(tdDiv);

            // Valores de tiempo
            timeCols.forEach(col => {
                const td = document.createElement('td');
                const val = parseFloat(row[col]);
                td.textContent = !isNaN(val) ? val.toLocaleString('es-PE', { maximumFractionDigits: 0 }) : '-';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    };

    appendRows(rRows, 'RECIBO');
    appendRows(dRows, 'DESPACHO');
    appendRows(iRows, 'INVENTARIO');
  }
});
