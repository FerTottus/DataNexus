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
      // 1. Obtener lista de pestañas para resolver nombres exactos (evita errores de espacios al final o mayúsculas)
      const tabs = await window.GoogleSheetsService.fetchSheetTabs(sheetId);
      const tabNames = tabs.map(t => t.title);
      
      const sheetGrafico = tabNames.find(t => t.trim().toUpperCase() === 'BD_GRAFICO' || t.trim().toUpperCase() === 'BD_GRÁFICO');
      const sheetBase = tabNames.find(t => {
          const name = t.trim().toUpperCase();
          return name === 'BASE' || name === 'BASES';
      });

      if (!sheetGrafico) {
          throw new Error(`No se encontró la hoja 'BD_Grafico'. Hojas disponibles: ${tabNames.join(', ')}`);
      }
      if (!sheetBase) {
          throw new Error(`No se encontró la hoja 'BASE' o 'BASES'. Hojas disponibles: ${tabNames.join(', ')}. Por favor verifica el nombre exacto en tu Google Sheet.`);
      }

      // Envolver en comillas simples por si el nombre tiene espacios
      const safeGrafico = `'${sheetGrafico}'`;
      const safeBase = `'${sheetBase}'`;

      // Data para los gráficos (Totales)
      window.graficoFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, `${safeGrafico}!A5:P250`);
      window.graficoSecos = await window.GoogleSheetsService.fetchSheetData(sheetId, `${safeGrafico}!R5:AH250`);

      // Data para las tablas de cajas (Divisiones)
      window.tablasFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, `${safeBase}!A5:L109`);
      window.tablasSecos = await window.GoogleSheetsService.fetchSheetData(sheetId, `${safeBase}!R5:AA109`);

      connectBox.classList.add('hidden');
      connectionSuccessInfo.classList.remove('hidden');
      document.getElementById('connectedSheetName').textContent = `Documento cargado (ID: ${sheetId.substring(0, 8)}...)`;
      dashboardSection.classList.remove('hidden');

      applyFiltersAndRender();
    } catch (e) {
      console.error(e);
      // Extraemos solo el mensaje para que sea legible
      alert("Error: " + e.message);
    } finally {
      loadingIndicator.classList.add('hidden');
    }
  });

  btnApplyFilters?.addEventListener('click', () => {
    applyFiltersAndRender();
  });

  function applyFiltersAndRender() {
    if(!window.graficoFrescos || !window.graficoSecos) return;
    const filterValue = document.getElementById('weeksFilter').value;
    
    processAndRenderSection('Secos', window.graficoSecos, window.tablasSecos, filterValue);
    processAndRenderSection('Frescos', window.graficoFrescos, window.tablasFrescos, filterValue);
  }

  // --- LÓGICA DE PROCESAMIENTO Y GRÁFICOS ---
  const charts = {};

  function processAndRenderSection(prefix, chartData, tableData, filterValue) {
    if(!chartData.rows || chartData.rows.length === 0) return;

    // 1. Detectar columna de fecha (e.g. "[34-2026]")
    let timeColName = null;
    for (let h of chartData.headers) {
      for(let i=0; i<Math.min(5, chartData.rows.length); i++) {
        let val = String(chartData.rows[i]?.[h] || '');
        if (val.includes('-202') || (val.includes('[') && val.includes(']'))) {
          timeColName = h;
          break;
        }
      }
      if(timeColName) break;
    }
    if (!timeColName) timeColName = chartData.headers[1] || chartData.headers[0];

    // Mapeo Flexible de Columnas de Métricas
    const getCol = (keywords) => chartData.headers.find(h => keywords.some(k => h.toUpperCase().includes(k)));
    const colInventario = getCol(['INVENTARIO ACT', 'INVENTARIO']);
    const colRecibo = getCol(['RECIBO', 'INGRESO']);
    const colDespacho = getCol(['DESPACHO', 'SALIDA']);
    const colPlanInv = getCol(['PLAN INV', 'PLAN INVENTARIO']);
    const colPlanRecibo = getCol(['PLAN RECIBO', 'PLAN INGRESO']);
    const colPlanDespacho = getCol(['PLAN DESPACHO', 'PLAN SALIDA']);

    // 2. Extraer filas válidas (que tengan fecha y no sean totales)
    let validRows = chartData.rows.filter(r => r[timeColName] && String(r[timeColName]).trim() !== '' && !String(r[timeColName]).toUpperCase().includes('TOTAL'));

    // 3. Determinar el "HOY" (La última fila que tenga datos reales, no ceros)
    // Buscamos desde el final hacia el principio hasta encontrar data > 0
    let lastDataIndex = validRows.length - 1;
    for(let i = validRows.length - 1; i >= 0; i--) {
        const row = validRows[i];
        const r = parseFloat(row[colRecibo]) || 0;
        const d = parseFloat(row[colDespacho]) || 0;
        const inv = parseFloat(row[colInventario]) || 0;
        if (r > 0 || d > 0 || inv > 0) {
            lastDataIndex = i;
            break;
        }
    }
    
    // Recortar filas futuras sin data
    validRows = validRows.slice(0, lastDataIndex + 1);
    let rowsToPlot = validRows;

    // 4. Aplicar Filtro "Últimas X Semanas" + Año Anterior secuencial
    if (filterValue !== 'all') {
        const numWeeks = parseInt(filterValue, 10);
        
        // El año actual lo sacamos de la última fila válida
        const lastValidRow = validRows[validRows.length - 1];
        const match = String(lastValidRow[timeColName]).match(/-(\d{4})/);
        
        if (match) {
            const currentYear = parseInt(match[1]);
            
            // Tomar las ultimas X semanas
            const lastWeeksCurrentYear = validRows.slice(-numWeeks);
            
            // Buscar equivalentes del año pasado
            const previousYearRows = [];
            lastWeeksCurrentYear.forEach(cr => {
                const prevLabel = String(cr[timeColName]).replace(`-${currentYear}`, `-${currentYear - 1}`);
                const pr = chartData.rows.find(r => String(r[timeColName]) === prevLabel);
                if (pr) previousYearRows.push(pr);
            });
            
            rowsToPlot = [...previousYearRows, ...lastWeeksCurrentYear];
        } else {
            rowsToPlot = validRows.slice(-numWeeks * 2);
        }
    }

    // Si la data está muy vacía, fallback
    if(rowsToPlot.length === 0) rowsToPlot = validRows;

    // Extraer vectores
    const labels = rowsToPlot.map(r => r[timeColName]);
    const inventario = rowsToPlot.map(r => parseFloat(r[colInventario]) || 0);
    const recibo = rowsToPlot.map(r => parseFloat(r[colRecibo]) || 0);
    const despacho = rowsToPlot.map(r => parseFloat(r[colDespacho]) || 0);
    const planInv = colPlanInv ? rowsToPlot.map(r => parseFloat(r[colPlanInv]) || 0) : [];
    const planRecibo = colPlanRecibo ? rowsToPlot.map(r => parseFloat(r[colPlanRecibo]) || 0) : [];
    const planDespacho = colPlanDespacho ? rowsToPlot.map(r => parseFloat(r[colPlanDespacho]) || 0) : [];

    renderMixedChart(prefix, labels, { inventario, recibo, despacho, planInv, planRecibo, planDespacho });
    
    // Renderizar la tabla base inferior
    renderBaseTables(prefix, tableData);
  }

  function renderMixedChart(prefix, labels, data) {
    const canvasId = `chart${prefix}`;
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    // Notas de Order: En Chart.js, mayor `order` significa que se dibuja POR DEBAJO de los menores.
    // Queremos:
    // Fondo: INVENTARIO (Area) -> order: 10
    // Medio: RECIBO y DESPACHO (Barras) -> order: 5
    // Frente: PLAN (Líneas punteadas) -> order: 1

    const datasets = [
      {
        type: 'line',
        label: 'INVENTARIO',
        data: data.inventario,
        backgroundColor: 'rgba(211, 211, 211, 0.7)', // Gris Área más sólido
        borderColor: '#9e9e9e',
        borderWidth: 2,
        fill: true,
        order: 10 
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
        order: 5
      }
    ];

    if(data.planInv && data.planInv.some(v => v > 0)) {
        datasets.push({
            type: 'line',
            label: 'PLAN INV',
            data: data.planInv,
            borderColor: '#2a9d8f', // Verde Plan
            borderWidth: 2,
            borderDash: [5, 5],
            fill: false,
            pointRadius: 0,
            order: 1
        });
    }

    if(data.planRecibo && data.planRecibo.some(v => v > 0)) {
        datasets.push({
            type: 'line',
            label: 'PLAN RECIBO',
            data: data.planRecibo,
            borderColor: '#1d3557', // Azul Oscuro Plan
            borderWidth: 2,
            borderDash: [5, 5],
            fill: false,
            pointRadius: 0,
            order: 1
        });
    }

    if(data.planDespacho && data.planDespacho.some(v => v > 0)) {
        datasets.push({
            type: 'line',
            label: 'PLAN DESPACHO',
            data: data.planDespacho,
            borderColor: '#9d0208', // Rojo Oscuro Plan
            borderWidth: 2,
            borderDash: [5, 5],
            fill: false,
            pointRadius: 0,
            order: 1
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
             grid: { display: false }
          },
          y: {
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

  function renderBaseTables(prefix, tableData) {
      // Como el sheet BASE tiene las 3 tablas de cajas (Recibo, Despacho, Inventario)
      // ya armadas y listas, simplemente vamos a escupir la data en HTML en su contenedor,
      // respetando su formato natural y evitando ensuciar con columnas que no vienen.
      
      const container = document.getElementById(`table${prefix}Container`);
      if (!container) return;
      
      if (!tableData || !tableData.rawValues || tableData.rawValues.length === 0) {
          container.innerHTML = '<p>No se encontraron datos en la hoja BASE.</p>';
          return;
      }

      let html = '<table class="data-table"><tbody>';
      
      tableData.rawValues.forEach(row => {
          let hasData = row.some(cell => String(cell).trim() !== '');
          if (!hasData) return; // Ignorar filas completamente vacías

          html += '<tr>';
          row.forEach(cell => {
              // Dar formato a números
              let text = cell;
              if(typeof cell === 'number') {
                  text = cell.toLocaleString('es-PE', { maximumFractionDigits: 0 });
              }
              
              // Si es un header de semana como "[34-2026]" o "Cajas (DIV)"
              let isHeader = String(cell).includes('Cajas') || String(cell).includes('-202');
              let tag = isHeader ? 'th' : 'td';
              
              html += `<${tag}>${text !== undefined && text !== null ? text : ''}</${tag}>`;
          });
          html += '</tr>';
      });

      html += '</tbody></table>';
      container.innerHTML = html;
  }
});
