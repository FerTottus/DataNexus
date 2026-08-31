document.addEventListener('DOMContentLoaded', () => {
  // Inicialización de Auth
  if (window.GoogleSheetsService) {
    window.GoogleSheetsService.initAuth();
    updateAuthUI();
  }

  // UI Elements
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

  // Explorador de Drive Modal
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
        // Auto cargar datos si se selecciona un archivo
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

  // Cambio de Pestañas (Tabs)
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

  // Cargar Datos Global
  btnFetchData?.addEventListener('click', async () => {
    const urlOrId = sheetUrlInput.value.trim();
    if (!urlOrId) return alert("Por favor, ingresa la URL o busca un archivo en tu Drive.");
    const sheetId = window.GoogleSheetsService.extractSpreadsheetId(urlOrId);
    if (!sheetId) return alert("ID inválido o URL no reconocida.");
    if (!window.GoogleSheetsService.isAuthenticated()) return alert("Debes iniciar sesión con Google primero.");

    loadingIndicator.classList.remove('hidden');
    
    try {
      // Ajustamos rangos más largos por si el usuario agregó más divisiones o semanas
      window.rawFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!A5:P200");
      window.rawSecos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!R5:AH200");

      connectBox.classList.add('hidden');
      connectionSuccessInfo.classList.remove('hidden');
      document.getElementById('connectedSheetName').textContent = `Documento cargado (ID: ${sheetId.substring(0, 8)}...)`;
      dashboardSection.classList.remove('hidden');

      applyFiltersAndRender();
    } catch (e) {
      console.error(e);
      alert("Error al cargar los datos: " + e.message);
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
    
    processAndRenderSection('Secos', window.rawSecos, filterType);
    processAndRenderSection('Frescos', window.rawFrescos, filterType);
  }

  // --- LÓGICA DE PROCESAMIENTO Y GRÁFICOS ---
  const charts = {};

  function processAndRenderSection(prefix, rawData, filterType) {
    if(!rawData.rows || rawData.rows.length === 0) return;

    const firstCol = rawData.headers[0];
    const timeCols = extractAndFilterTimeCols(rawData.headers, filterType);

    // Separar datos en bloques según el Dashboard de ejemplo
    const blocks = { RECIBO: [], DESPACHO: [], INVENTARIO: [] };
    let currentBlock = 'RECIBO'; // por defecto el primero

    rawData.rows.forEach(row => {
      const label = String(row[firstCol] || '').toUpperCase();
      if(label.includes('RECIBO') || label.includes('INGRESO')) currentBlock = 'RECIBO';
      else if(label.includes('DESPACHO') || label.includes('SALIDA')) currentBlock = 'DESPACHO';
      else if(label.includes('INVENTARIO')) currentBlock = 'INVENTARIO';
      else if(label && !label.includes('TOTAL') && !label.includes('CAJAS') && !label.includes('DIVISION')) {
        blocks[currentBlock].push(row);
      }
    });

    // Calcular Totales por Columna para el gráfico principal
    const totals = { RECIBO: [], DESPACHO: [], INVENTARIO: [] };
    
    ['RECIBO', 'DESPACHO', 'INVENTARIO'].forEach(blockType => {
        timeCols.forEach(col => {
            let sum = 0;
            blocks[blockType].forEach(row => {
                sum += parseFloat(row[col]) || 0;
            });
            totals[blockType].push(sum);
        });
    });

    // Renderizar Tablas individuales
    renderTableBlock(`${prefix}Recibo`, timeCols, blocks.RECIBO, firstCol, totals.RECIBO);
    renderTableBlock(`${prefix}Despacho`, timeCols, blocks.DESPACHO, firstCol, totals.DESPACHO);
    renderTableBlock(`${prefix}Inventario`, timeCols, blocks.INVENTARIO, firstCol, totals.INVENTARIO);

    // Renderizar un único Gráfico Combinado
    renderComboChart(prefix, timeCols, totals);
  }

  function extractAndFilterTimeCols(headers, filterType) {
    let cols = headers.slice(1).filter(h => !h.toUpperCase().includes('TOTAL') && h.trim() !== '');
    
    if (filterType === 'all') return cols;

    // "Ultimos 2 meses" -> Tomamos aproximadamente 8 o 9 semanas recientes
    if (filterType === 'last2months') {
        return cols.slice(-9); 
    }

    if (filterType === 'last2months_yoy') {
        const recentCols = cols.slice(-9);
        const lastYearCols = [];
        
        recentCols.forEach(col => {
            // Ejemplo: "[6-2026]" -> "[6-2025]"
            let match = col.match(/(\d{4})/);
            if (match) {
                let year = parseInt(match[1]);
                let prevYearCol = col.replace(year.toString(), (year-1).toString());
                if(cols.includes(prevYearCol)) lastYearCols.push(prevYearCol);
            }
        });
        
        return lastYearCols.length > 0 ? lastYearCols : cols.slice(0, 9); 
    }

    return cols;
  }

  function renderTableBlock(tableId, timeCols, rows, labelCol, totalsArr) {
    const thead = document.getElementById(`table${tableId}Header`);
    const tbody = document.getElementById(`table${tableId}Body`);
    const tfoot = document.getElementById(`table${tableId}Foot`);
    
    if(!thead || !tbody || !tfoot) return;

    thead.innerHTML = '';
    tbody.innerHTML = '';
    tfoot.innerHTML = '';

    // Headers
    const thEmpty = document.createElement('th');
    thEmpty.textContent = ''; 
    thead.appendChild(thEmpty);
    
    timeCols.forEach(col => {
        const th = document.createElement('th');
        th.className = 'table-header-color';
        th.textContent = col;
        thead.appendChild(th);
    });

    // Rows
    rows.forEach(row => {
        const tr = document.createElement('tr');
        
        // Division
        const tdDiv = document.createElement('td');
        tdDiv.textContent = row[labelCol] || '';
        tr.appendChild(tdDiv);

        // Valores
        timeCols.forEach(col => {
            const td = document.createElement('td');
            const val = parseFloat(row[col]);
            td.textContent = !isNaN(val) ? val.toLocaleString('es-PE', { maximumFractionDigits: 0 }) : '-';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    // Footer (Total)
    const trFoot = document.createElement('tr');
    trFoot.className = 'total-row';
    const tdTotalLabel = document.createElement('td');
    tdTotalLabel.textContent = 'Total';
    trFoot.appendChild(tdTotalLabel);

    totalsArr.forEach(totalVal => {
        const td = document.createElement('td');
        td.textContent = totalVal.toLocaleString('es-PE', { maximumFractionDigits: 0 });
        trFoot.appendChild(td);
    });
    tfoot.appendChild(trFoot);
  }

  function renderComboChart(prefix, xLabels, totals) {
    const canvasId = `chart${prefix}Combined`;
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    // Recreamos exactamente el gráfico de los mockups (Inventario como área, Recibo y Despacho como barras)
    // El orden importa para la superposición: área primero (Inventario)
    
    const datasets = [
      {
        type: 'line',
        label: 'INVENTARIO',
        data: totals.INVENTARIO,
        backgroundColor: 'rgba(169, 169, 169, 0.4)', // Gris semi-transparente
        borderColor: '#7a7a7a',
        borderWidth: 2,
        fill: true,
        tension: 0.3,
        yAxisID: 'y'
      },
      {
        type: 'bar',
        label: 'RECIBO',
        data: totals.RECIBO,
        backgroundColor: '#457b9d', // Azul
        borderColor: '#1d3557',
        borderWidth: 1,
        yAxisID: 'y'
      },
      {
        type: 'bar',
        label: 'DESPACHO',
        data: totals.DESPACHO,
        backgroundColor: '#e76f51', // Naranja/Rojo
        borderColor: '#d00000',
        borderWidth: 1,
        yAxisID: 'y'
      }
    ];

    charts[canvasId] = new Chart(ctx, {
      data: {
        labels: xLabels,
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
               // Formatear números grandes con comas
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
});
