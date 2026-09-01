document.addEventListener('DOMContentLoaded', () => {
  // ══════════════════════════════════════════════
  // AUTH
  // ══════════════════════════════════════════════
  if (window.GoogleSheetsService) {
    window.GoogleSheetsService.initAuth();
    updateAuthUI();
  }

  function updateAuthUI() {
    const isAuth = window.GoogleSheetsService.isAuthenticated();
    document.getElementById('btnGoogleLogin')?.classList.toggle('hidden', isAuth);
    document.getElementById('userInfo')?.classList.toggle('hidden', !isAuth);
    if (isAuth && window.GoogleSheetsService.userEmail) {
      document.getElementById('userEmail').textContent = window.GoogleSheetsService.userEmail;
    }
  }

  document.getElementById('btnGoogleLogin')?.addEventListener('click', () => {
    window.GoogleSheetsService.requestAccessToken((ok) => { if (ok) updateAuthUI(); });
  });
  document.getElementById('btnGoogleLogout')?.addEventListener('click', () => {
    window.GoogleSheetsService.logout(); updateAuthUI();
  });

  // ══════════════════════════════════════════════
  // DRIVE MODAL
  // ══════════════════════════════════════════════
  document.getElementById('btnBrowseDrive')?.addEventListener('click', () => {
    if (!window.GoogleSheetsService.isAuthenticated()) {
      window.GoogleSheetsService.requestAccessToken((ok) => { if (ok) { updateAuthUI(); openDriveModal(); } });
    } else { openDriveModal(); }
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
    fileList.innerHTML = ''; loading.classList.remove('hidden'); emptyMsg.classList.add('hidden');
    const files = await window.GoogleSheetsService.fetchRecentSpreadsheets();
    loading.classList.add('hidden');
    if (!files || files.length === 0) { emptyMsg.classList.remove('hidden'); return; }
    window._driveFiles = files;
    renderDriveFiles(files);
  }

  function renderDriveFiles(files) {
    const fileList = document.getElementById('drive-file-list');
    fileList.innerHTML = '';
    files.forEach(file => {
      const li = document.createElement('li');
      li.className = 'file-item'; li.style.cursor = 'pointer'; li.style.padding = '10px'; li.style.borderBottom = '1px solid #eee';
      li.innerHTML = `<i class="fa-solid fa-file-excel text-success"></i> <span style="margin-left:8px;">${file.name}</span>`;
      li.addEventListener('click', () => {
        document.getElementById('sheetUrlInput').value = file.id;
        document.getElementById('drive-modal').classList.add('hidden');
        document.getElementById('btnFetchData').click();
      });
      fileList.appendChild(li);
    });
  }
  window.filterDriveList = function(text) {
    if (!window._driveFiles) return;
    renderDriveFiles(window._driveFiles.filter(f => f.name.toLowerCase().includes(text.toLowerCase())));
  };

  // ══════════════════════════════════════════════
  // TABS
  // ══════════════════════════════════════════════
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.target).classList.add('active');
    });
  });
  document.getElementById('btnChangeSheet')?.addEventListener('click', () => {
    document.getElementById('connectBox').classList.remove('hidden');
    document.getElementById('connectionSuccessInfo').classList.add('hidden');
  });

  // ══════════════════════════════════════════════
  // FETCH DATA
  // ══════════════════════════════════════════════
  document.getElementById('btnFetchData')?.addEventListener('click', async () => {
    const urlOrId = document.getElementById('sheetUrlInput').value.trim();
    if (!urlOrId) return alert("Ingresa la URL o busca un archivo.");
    const sheetId = window.GoogleSheetsService.extractSpreadsheetId(urlOrId);
    if (!sheetId) return alert("ID inválido.");
    if (!window.GoogleSheetsService.isAuthenticated()) return alert("Inicia sesión primero.");

    document.getElementById('loadingIndicator').classList.remove('hidden');

    try {
      // Resolver nombre exacto de BD_Grafico
      const tabs = await window.GoogleSheetsService.fetchSheetTabs(sheetId);
      const tabNames = tabs.map(t => t.title);
      const sheetGrafico = tabNames.find(t => t.trim().toUpperCase().replace('Á','A') === 'BD_GRAFICO');
      if (!sheetGrafico) throw new Error(`No se encontró 'BD_Grafico'. Hojas: ${tabNames.join(', ')}`);
      const sg = `'${sheetGrafico}'`;

      // Rangos CORRECTOS según estructura real del archivo
      // Frescos: Columnas A-L, Filas 5-109
      // Secos:   Columnas V-AE, Filas 5-109
      window.dataFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, `${sg}!A5:L109`);
      window.dataSecos   = await window.GoogleSheetsService.fetchSheetData(sheetId, `${sg}!V5:AE109`);

      document.getElementById('connectBox').classList.add('hidden');
      document.getElementById('connectionSuccessInfo').classList.remove('hidden');
      document.getElementById('connectedSheetName').textContent = `ID: ${sheetId.substring(0, 10)}...`;
      document.getElementById('dashboardSection').classList.remove('hidden');

      renderAll();
    } catch (e) {
      console.error(e);
      alert("Error: " + e.message);
    } finally {
      document.getElementById('loadingIndicator').classList.add('hidden');
    }
  });

  document.getElementById('btnApplyFilters')?.addEventListener('click', renderAll);

  function renderAll() {
    if (!window.dataFrescos || !window.dataSecos) return;
    const numWeeks = document.getElementById('weeksFilter').value;

    // Frescos: col A = Semana (time key)
    renderSection('Frescos', window.dataFrescos, resolveColumns(window.dataFrescos.headers, window.dataFrescos.rows), numWeeks);

    // Secos: col W = SEMANA (time key)
    renderSection('Secos', window.dataSecos, resolveColumns(window.dataSecos.headers, window.dataSecos.rows), numWeeks);
  }

  /**
   * Busca dinámicamente los encabezados por coincidencia parcial (case insensitive, trimmed).
   * Para timeCol, revisa los datos reales para encontrar la columna con formato [X-YYYY].
   */
  function resolveColumns(headers, rows) {
    console.log('Headers recibidos:', JSON.stringify(headers));

    const find = (keywords) => {
      // Match exacto primero
      for (const h of headers) {
        const clean = h.trim().toUpperCase();
        for (const kw of keywords) {
          if (clean === kw) return h;
        }
      }
      // Match parcial
      for (const h of headers) {
        const clean = h.trim().toUpperCase();
        for (const kw of keywords) {
          if (clean.includes(kw)) return h;
        }
      }
      return null;
    };

    // Para timeCol: buscar en los datos reales cuál columna tiene formato [X-YYYY]
    let timeCol = null;
    if (rows && rows.length > 0) {
      for (const h of headers) {
        const sampleVal = String(rows[0][h] || '');
        if (/^\[?\d{1,2}-\d{4}\]?$/.test(sampleVal.trim())) {
          timeCol = h;
          break;
        }
      }
    }
    // Fallback por nombre
    if (!timeCol) timeCol = find(['SEMANA']) || headers[0];

    const recibo = find(['RECIBO']);
    const despacho = find(['DESPACHO']);
    const inventario = find(['INVENTARIO ACT']) || find(['INVENTARIO']);
    const planRecibo = find(['PLAN RECIBO']);
    const planDespacho = find(['PLAN DESPACHO']);
    const planInv = find(['PLAN INV']);

    const resolved = { timeCol, recibo, despacho, inventario, planRecibo, planDespacho, planInv };
    console.log('Columnas resueltas:', resolved);
    return resolved;
  }

  // ══════════════════════════════════════════════
  // CORE LOGIC
  // ══════════════════════════════════════════════
  const charts = {};

  /**
   * Parsea "[34-2026]" → { week: 34, year: 2026 }
   */
  function parseWeekLabel(label) {
    const s = String(label).replace(/[\[\]]/g, '');
    const parts = s.split('-');
    if (parts.length !== 2) return null;
    return { week: parseInt(parts[0], 10), year: parseInt(parts[1], 10) };
  }

  function renderSection(prefix, data, cols, filterValue) {
    if (!data.rows || data.rows.length === 0) return;

    // 1. Parsear todas las filas con su semana/año
    const allParsed = [];
    data.rows.forEach(row => {
      const label = row[cols.timeCol];
      if (!label) return;
      const parsed = parseWeekLabel(label);
      if (!parsed || isNaN(parsed.week) || isNaN(parsed.year)) return;

      allParsed.push({
        label,
        week: parsed.week,
        year: parsed.year,
        recibo:       parseFloat(row[cols.recibo]) || 0,
        despacho:     parseFloat(row[cols.despacho]) || 0,
        inventario:   parseFloat(row[cols.inventario]) || 0,
        planRecibo:   parseFloat(row[cols.planRecibo]) || 0,
        planDespacho: parseFloat(row[cols.planDespacho]) || 0,
        planInv:      parseFloat(row[cols.planInv]) || 0
      });
    });

    if (allParsed.length === 0) return;

    // 2. Determinar año actual = el mayor año en los datos
    const currentYear = Math.max(...allParsed.map(r => r.year));
    const prevYear = currentYear - 1;

    // 3. Encontrar la última semana con datos reales en el año actual
    const currentYearRows = allParsed
      .filter(r => r.year === currentYear)
      .sort((a, b) => a.week - b.week);

    let lastDataWeek = 0;
    for (let i = currentYearRows.length - 1; i >= 0; i--) {
      const r = currentYearRows[i];
      if (r.recibo > 0 || r.despacho > 0 || r.inventario > 0) {
        lastDataWeek = r.week;
        break;
      }
    }
    if (lastDataWeek === 0 && currentYearRows.length > 0) {
      lastDataWeek = currentYearRows[currentYearRows.length - 1].week;
    }

    // 4. Determinar rango de semanas a mostrar
    let weekNumbers;
    if (filterValue === 'all') {
      // Mostrar todas las semanas que tengan data en el año actual
      weekNumbers = currentYearRows
        .filter(r => r.recibo > 0 || r.despacho > 0 || r.inventario > 0)
        .map(r => r.week);
    } else {
      const n = parseInt(filterValue, 10);
      const startWeek = Math.max(1, lastDataWeek - n + 1);
      weekNumbers = [];
      for (let w = startWeek; w <= lastDataWeek; w++) weekNumbers.push(w);
    }

    if (weekNumbers.length === 0) return;

    // 5. Indexar datos por (año, semana) para acceso rápido
    const dataMap = {};
    allParsed.forEach(r => {
      dataMap[`${r.year}-${r.week}`] = r;
    });

    // 6. Construir arrays alineados por semana
    const labels = weekNumbers.map(w => `S${w}`);
    const get = (year, week, field) => {
      const key = `${year}-${week}`;
      return dataMap[key] ? dataMap[key][field] : 0;
    };

    const prevRecibo =       weekNumbers.map(w => get(prevYear, w, 'recibo'));
    const currRecibo =       weekNumbers.map(w => get(currentYear, w, 'recibo'));
    const planRecibo =       weekNumbers.map(w => get(currentYear, w, 'planRecibo'));

    const prevDespacho =     weekNumbers.map(w => get(prevYear, w, 'despacho'));
    const currDespacho =     weekNumbers.map(w => get(currentYear, w, 'despacho'));
    const planDespacho =     weekNumbers.map(w => get(currentYear, w, 'planDespacho'));

    const prevInventario =   weekNumbers.map(w => get(prevYear, w, 'inventario'));
    const currInventario =   weekNumbers.map(w => get(currentYear, w, 'inventario'));
    const planInv =          weekNumbers.map(w => get(currentYear, w, 'planInv'));

    // 7. Renderizar los 3 gráficos
    renderBarChart(`chart${prefix}Recibo`, '📦 Recibo', labels, [
      { label: `Recibo ${prevYear}`, data: prevRecibo, bg: 'rgba(100,149,237,0.5)', border: '#6495ED' },
      { label: `Recibo ${currentYear}`, data: currRecibo, bg: '#1d3557', border: '#1d3557' },
    ], planRecibo.some(v => v > 0) ? { label: 'PLAN RECIBO', data: planRecibo, color: '#457b9d' } : null);

    renderBarChart(`chart${prefix}Despacho`, '🚛 Despacho', labels, [
      { label: `Despacho ${prevYear}`, data: prevDespacho, bg: 'rgba(233,150,122,0.6)', border: '#E9967A' },
      { label: `Despacho ${currentYear}`, data: currDespacho, bg: '#e76f51', border: '#e76f51' },
    ], planDespacho.some(v => v > 0) ? { label: 'PLAN DESPACHO', data: planDespacho, color: '#9d0208' } : null);

    renderAreaChart(`chart${prefix}Inventario`, '📊 Inventario', labels, [
      { label: `Inventario ${prevYear}`, data: prevInventario, bg: 'rgba(200,200,200,0.4)', border: '#bbb' },
      { label: `Inventario ${currentYear}`, data: currInventario, bg: 'rgba(100,100,100,0.3)', border: '#555' },
    ], planInv.some(v => v > 0) ? { label: 'PLAN INV', data: planInv, color: '#2a9d8f' } : null);

    // 8. Tabla resumen (últimas 7 semanas del año actual)
    const tableWeeks = weekNumbers.slice(-7);
    renderTable(prefix, tableWeeks, currentYear, dataMap);
  }

  // ══════════════════════════════════════════════
  // CHART RENDERERS
  // ══════════════════════════════════════════════
  function renderBarChart(canvasId, title, labels, barSeries, planLine) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    const datasets = barSeries.map(s => ({
      type: 'bar',
      label: s.label,
      data: s.data,
      backgroundColor: s.bg,
      borderColor: s.border,
      borderWidth: 1
    }));

    if (planLine) {
      datasets.push({
        type: 'line',
        label: planLine.label,
        data: planLine.data,
        borderColor: planLine.color,
        borderWidth: 2,
        borderDash: [6, 3],
        fill: false,
        pointRadius: 0,
        tension: 0.2
      });
    }

    charts[canvasId] = new Chart(ctx, {
      data: { labels, datasets },
      options: chartOptions()
    });
  }

  function renderAreaChart(canvasId, title, labels, areaSeries, planLine) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    const datasets = areaSeries.map(s => ({
      type: 'line',
      label: s.label,
      data: s.data,
      backgroundColor: s.bg,
      borderColor: s.border,
      borderWidth: 2,
      fill: true,
      tension: 0.3,
      pointRadius: 2
    }));

    if (planLine) {
      datasets.push({
        type: 'line',
        label: planLine.label,
        data: planLine.data,
        borderColor: planLine.color,
        borderWidth: 2,
        borderDash: [6, 3],
        fill: false,
        pointRadius: 0,
        tension: 0.2
      });
    }

    charts[canvasId] = new Chart(ctx, {
      data: { labels, datasets },
      options: chartOptions()
    });
  }

  function chartOptions() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { usePointStyle: true, boxWidth: 10, font: { size: 11 } } },
        tooltip: { mode: 'index', intersect: false }
      },
      scales: {
        x: { grid: { display: false } },
        y: {
          beginAtZero: true,
          ticks: { callback: v => v.toLocaleString('es-PE') }
        }
      },
      interaction: { mode: 'index', intersect: false }
    };
  }

  // ══════════════════════════════════════════════
  // TABLE RENDERER
  // ══════════════════════════════════════════════
  function renderTable(prefix, weekNumbers, year, dataMap) {
    const container = document.getElementById(`table${prefix}Container`);
    if (!container) return;

    let html = `<div class="table-wrapper-title">Resumen últimas 7 semanas (${year})</div>`;
    html += '<table class="data-table"><thead><tr><th>Semana</th>';
    weekNumbers.forEach(w => { html += `<th>[${w}-${year}]</th>`; });
    html += '</tr></thead><tbody>';

    ['recibo', 'despacho', 'inventario'].forEach(metric => {
      const label = metric === 'recibo' ? 'RECIBO' : metric === 'despacho' ? 'DESPACHO' : 'INVENTARIO';
      html += `<tr><td><strong>${label}</strong></td>`;
      weekNumbers.forEach(w => {
        const val = dataMap[`${year}-${w}`] ? dataMap[`${year}-${w}`][metric] : 0;
        html += `<td>${val.toLocaleString('es-PE', { maximumFractionDigits: 0 })}</td>`;
      });
      html += '</tr>';
    });

    html += '</tbody></table>';
    container.innerHTML = html;
  }
});
