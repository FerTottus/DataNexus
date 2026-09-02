document.addEventListener('DOMContentLoaded', () => {
  // ══════════════════════════════════════════════
  // CHART DATA LABELS PLUGIN REGISTRATION
  // ══════════════════════════════════════════════
  if (window.ChartDataLabels) {
    Chart.register(window.ChartDataLabels);
  }

  // ══════════════════════════════════════════════
  // AUTH INITIALIZATION
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
      li.innerHTML = `<i class="fa-solid fa-file-excel text-success"></i> <span style="margin-left:8px; font-weight:600;">${file.name}</span>`;
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
  // TABS & INTERACTION
  // ══════════════════════════════════════════════
  document.querySelectorAll('.glass-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.glass-tab-btn').forEach(b => b.classList.remove('active'));
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
    if (!urlOrId) return alert("Ingresa la URL o busca un archivo en Google Drive.");
    const sheetId = window.GoogleSheetsService.extractSpreadsheetId(urlOrId);
    if (!sheetId) return alert("ID de hoja de cálculo inválido.");
    if (!window.GoogleSheetsService.isAuthenticated()) return alert("Inicia sesión primero.");

    document.getElementById('loadingIndicator').classList.remove('hidden');

    try {
      // Resolver nombre exacto de la pestaña BD_Grafico
      const tabs = await window.GoogleSheetsService.fetchSheetTabs(sheetId);
      const tabNames = tabs.map(t => t.title);
      const sheetGrafico = tabNames.find(t => t.trim().toUpperCase().replace('Á','A') === 'BD_GRAFICO');
      if (!sheetGrafico) throw new Error(`No se encontró 'BD_Grafico'. Pestañas disponibles: ${tabNames.join(', ')}`);
      const sg = `'${sheetGrafico}'`;

      // ─────────────────────────────────────────────
      // RANGOS PRECISOS DE LA HOJA BD_Grafico:
      // Frescos: Columnas A a L (A5:L109)
      // Secos:   Columnas V a AE (V5:AE109) -> Incluye PLAN DESPACHO y PLAN INV
      // ─────────────────────────────────────────────
      window.dataFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, `${sg}!A5:L109`);
      window.dataSecos   = await window.GoogleSheetsService.fetchSheetData(sheetId, `${sg}!V5:AE109`);

      document.getElementById('connectBox').classList.add('hidden');
      document.getElementById('connectionSuccessInfo').classList.remove('hidden');
      document.getElementById('connectedSheetName').textContent = `Documento cargado (ID: ${sheetId.substring(0, 12)}...)`;
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
  document.getElementById('toggleDataLabels')?.addEventListener('change', renderAll);

  function renderAll() {
    if (!window.dataFrescos || !window.dataSecos) return;
    const numWeeks = document.getElementById('weeksFilter').value;
    const showLabels = document.getElementById('toggleDataLabels')?.checked ?? true;

    // CD Frescos
    renderSection('Frescos', window.dataFrescos, resolveColumns(window.dataFrescos.headers, window.dataFrescos.rows), numWeeks, showLabels);

    // CD Secos
    renderSection('Secos', window.dataSecos, resolveColumns(window.dataSecos.headers, window.dataSecos.rows), numWeeks, showLabels);
  }

  // ══════════════════════════════════════════════
  // DYNAMIC COLUMN RESOLUTION
  // ══════════════════════════════════════════════
  function resolveColumns(headers, rows) {
    const find = (keywords) => {
      // 1. Match exacto
      for (const h of headers) {
        const clean = h.trim().toUpperCase();
        for (const kw of keywords) {
          if (clean === kw) return h;
        }
      }
      // 2. Match parcial
      for (const h of headers) {
        const clean = h.trim().toUpperCase();
        for (const kw of keywords) {
          if (clean.includes(kw)) return h;
        }
      }
      return null;
    };

    // Detectar columna de fecha en base al contenido [X-YYYY]
    let timeCol = null;
    if (rows && rows.length > 0) {
      for (const h of headers) {
        for (let i = 0; i < Math.min(5, rows.length); i++) {
          const val = String(rows[i][h] || '');
          if (/^\[?\d{1,2}-\d{4}\]?$/.test(val.trim())) {
            timeCol = h;
            break;
          }
        }
        if (timeCol) break;
      }
    }
    if (!timeCol) timeCol = find(['SEMANA', 'SEMANA2', 'SEMAÑO']) || headers[0];

    const recibo = find(['RECIBO']);
    const despacho = find(['DESPACHO']);
    const inventario = find(['INVENTARIO ACT']) || find(['INVENTARIO']);
    const planRecibo = find(['PLAN RECIBO', 'PLAN_RECIBO']);
    const planDespacho = find(['PLAN DESPACHO', 'PLAN_DESPACHO']);
    const planInv = find(['PLAN INV', 'PLAN_INV', 'PLAN INVENTARIO']);

    return { timeCol, recibo, despacho, inventario, planRecibo, planDespacho, planInv };
  }

  // ══════════════════════════════════════════════
  // CORE PROCESSING & SECTION RENDERER
  // ══════════════════════════════════════════════
  const charts = {};

  function parseWeekLabel(label) {
    const s = String(label).replace(/[\[\]]/g, '');
    const parts = s.split('-');
    if (parts.length !== 2) return null;
    return { week: parseInt(parts[0], 10), year: parseInt(parts[1], 10) };
  }

  function renderSection(prefix, data, cols, filterValue, showLabels) {
    if (!data.rows || data.rows.length === 0) return;

    // 1. Extraer y estructurar datos
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

    // 2. Determinar año actual (el más reciente con data)
    const currentYear = Math.max(...allParsed.map(r => r.year));
    const prevYear = currentYear - 1;

    // 3. Encontrar última semana con movimientos reales en el año actual
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

    // 4. Semanas a graficar
    let weekNumbers;
    if (filterValue === 'all') {
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

    // 5. Mapeo de datos por año y semana
    const dataMap = {};
    allParsed.forEach(r => {
      dataMap[`${r.year}-${r.week}`] = r;
    });

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

    // 6. Renderizar KPI Executive Strip
    renderKPIs(prefix, weekNumbers, currentYear, prevYear, dataMap);

    // 7. Renderizar Gráficos con Paleta Vibrante Glass y DataLabels
    // 📦 RECIBO: Indigo / Blue Palette + Cyan Plan
    renderBarChart(`chart${prefix}Recibo`, labels, [
      { label: `Recibo ${prevYear}`, data: prevRecibo, bg: 'rgba(147, 197, 253, 0.75)', border: '#60a5fa' },
      { label: `Recibo ${currentYear}`, data: currRecibo, bg: '#2563eb', border: '#1d4ed8' },
    ], planRecibo.some(v => v > 0) ? { label: 'PLAN RECIBO', data: planRecibo, color: '#0284c7' } : null, showLabels);

    // 🚛 DESPACHO: Coral / Flame Orange Palette + Ruby Plan
    renderBarChart(`chart${prefix}Despacho`, labels, [
      { label: `Despacho ${prevYear}`, data: prevDespacho, bg: 'rgba(253, 186, 116, 0.75)', border: '#fb923c' },
      { label: `Despacho ${currentYear}`, data: currDespacho, bg: '#ea580c', border: '#c2410c' },
    ], planDespacho.some(v => v > 0) ? { label: 'PLAN DESPACHO', data: planDespacho, color: '#e11d48' } : null, showLabels);

    // 📊 INVENTARIO: Frosted Emerald Glass Area + Mint Plan
    renderAreaChart(`chart${prefix}Inventario`, labels, [
      { label: `Inventario ${prevYear}`, data: prevInventario, bg: 'rgba(148, 163, 184, 0.25)', border: '#94a3b8' },
      { label: `Inventario ${currentYear}`, data: currInventario, bg: 'rgba(16, 185, 129, 0.35)', border: '#059669' },
    ], planInv.some(v => v > 0) ? { label: 'PLAN INV', data: planInv, color: '#0d9488' } : null, showLabels);

    // 8. Tabla Resumen (Últimas semanas seleccionadas)
    renderTable(prefix, weekNumbers.slice(-7), currentYear, prevYear, dataMap);
  }

  // ══════════════════════════════════════════════
  // KPI EXECUTIVE SUMMARY RENDERER
  // ══════════════════════════════════════════════
  function renderKPIs(prefix, weekNumbers, currentYear, prevYear, dataMap) {
    const container = document.getElementById(`kpi${prefix}Container`);
    if (!container) return;

    let totReciboCurr = 0, totReciboPrev = 0, totPlanRecibo = 0;
    let totDespachoCurr = 0, totDespachoPrev = 0, totPlanDespacho = 0;
    let totInvCurr = 0, totInvPrev = 0, countInv = 0;

    weekNumbers.forEach(w => {
      const c = dataMap[`${currentYear}-${w}`];
      const p = dataMap[`${prevYear}-${w}`];
      if (c) {
        totReciboCurr += c.recibo || 0;
        totPlanRecibo += c.planRecibo || 0;
        totDespachoCurr += c.despacho || 0;
        totPlanDespacho += c.planDespacho || 0;
        if (c.inventario > 0) {
          totInvCurr += c.inventario;
          countInv++;
        }
      }
      if (p) {
        totReciboPrev += p.recibo || 0;
        totDespachoPrev += p.despacho || 0;
        totInvPrev += p.inventario || 0;
      }
    });

    const avgInvCurr = countInv > 0 ? totInvCurr / countInv : 0;
    const avgInvPrev = weekNumbers.length > 0 ? totInvPrev / weekNumbers.length : 0;

    const calcYoY = (curr, prev) => {
      if (prev === 0) return { pct: '0.0%', isUp: true };
      const diff = ((curr - prev) / prev) * 100;
      return {
        pct: (diff >= 0 ? '+' : '') + diff.toFixed(1) + '%',
        isUp: diff >= 0
      };
    };

    const yoyRecibo = calcYoY(totReciboCurr, totReciboPrev);
    const yoyDespacho = calcYoY(totDespachoCurr, totDespachoPrev);
    const yoyInv = calcYoY(avgInvCurr, avgInvPrev);

    const fmt = n => Math.round(n).toLocaleString('es-PE');

    container.innerHTML = `
      <div class="kpi-card kpi-recibo">
        <div class="kpi-title"><i class="fa-solid fa-boxes-packing text-primary"></i> Total Entradas (Recibo)</div>
        <div class="kpi-value">${fmt(totReciboCurr)}</div>
        <div class="kpi-sub">
          <span class="kpi-badge ${yoyRecibo.isUp ? 'badge-up' : 'badge-down'}">
            <i class="fa-solid fa-arrow-${yoyRecibo.isUp ? 'trend-up' : 'trend-down'}"></i> ${yoyRecibo.pct} YoY
          </span>
          <span>vs ${fmt(totReciboPrev)} (${prevYear})</span>
        </div>
      </div>

      <div class="kpi-card kpi-despacho">
        <div class="kpi-title"><i class="fa-solid fa-truck-fast" style="color:#ea580c;"></i> Total Salidas (Despacho)</div>
        <div class="kpi-value">${fmt(totDespachoCurr)}</div>
        <div class="kpi-sub">
          <span class="kpi-badge ${yoyDespacho.isUp ? 'badge-up' : 'badge-down'}">
            <i class="fa-solid fa-arrow-${yoyDespacho.isUp ? 'trend-up' : 'trend-down'}"></i> ${yoyDespacho.pct} YoY
          </span>
          <span>vs ${fmt(totDespachoPrev)} (${prevYear})</span>
        </div>
      </div>

      <div class="kpi-card kpi-inventario">
        <div class="kpi-title"><i class="fa-solid fa-warehouse" style="color:#059669;"></i> Stock Inventario (Promedio)</div>
        <div class="kpi-value">${fmt(avgInvCurr)}</div>
        <div class="kpi-sub">
          <span class="kpi-badge ${yoyInv.isUp ? 'badge-up' : 'badge-down'}">
            <i class="fa-solid fa-arrow-${yoyInv.isUp ? 'trend-up' : 'trend-down'}"></i> ${yoyInv.pct} YoY
          </span>
          <span>vs ${fmt(avgInvPrev)} (${prevYear})</span>
        </div>
      </div>
    `;
  }

  // ══════════════════════════════════════════════
  // CHART BUILDERS WITH DATALABELS & LIQUID THEME
  // ══════════════════════════════════════════════
  function formatNumberBadge(val) {
    if (!val || val === 0) return '';
    if (val >= 1000000) return (val / 1000000).toFixed(2) + 'M';
    if (val >= 1000) return Math.round(val / 1000).toLocaleString('es-PE') + 'k';
    return Math.round(val).toLocaleString('es-PE');
  }

  function getBaseChartOptions(showLabels) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: { top: showLabels ? 34 : 12, bottom: 5, left: 12, right: 12 }
      },
      plugins: {
        legend: {
          position: 'top',
          align: 'center', // Leyenda centrada
          labels: {
            usePointStyle: true,
            boxWidth: 9,
            padding: 18,
            font: { size: 11.5, weight: '600', family: 'Inter' },
            color: '#334155'
          }
        },
        tooltip: {
          mode: 'index',
          intersect: false,
          backgroundColor: 'rgba(15, 23, 42, 0.92)',
          titleFont: { size: 12, weight: '700' },
          bodyFont: { size: 11 },
          padding: 10,
          cornerRadius: 8,
          callbacks: {
            label: (ctx) => {
              const label = ctx.dataset.label || '';
              const val = ctx.parsed.y !== null ? Math.round(ctx.parsed.y).toLocaleString('es-PE') : '0';
              return `  ${label}: ${val} cajas`;
            }
          }
        },
        datalabels: {
          // Las opciones específicas por dataset tienen prioridad
          display: showLabels,
          clip: false,
          clamp: true
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { font: { weight: '600', size: 11, family: 'Inter' }, color: '#475569' }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(226, 232, 240, 0.6)' },
          ticks: {
            font: { size: 10.5, family: 'Inter' },
            color: '#64748b',
            callback: (v) => formatNumberBadge(v)
          }
        }
      },
      interaction: { mode: 'index', intersect: false }
    };
  }

  function renderBarChart(canvasId, labels, barSeries, planLine, showLabels) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    // barSeries[0] = 2025 (Año Anterior)
    // barSeries[1] = 2026 (Año Actual)
    const datasets = [
      {
        type: 'bar',
        label: barSeries[0].label,
        data: barSeries[0].data,
        backgroundColor: barSeries[0].bg,
        borderColor: barSeries[0].border,
        borderWidth: 1.5,
        borderRadius: 6,
        barPercentage: 0.75,
        categoryPercentage: 0.8,
        datalabels: {
          display: showLabels,
          // Año anterior: colocado en el centro de la barra para evitar colisiones superiores
          anchor: 'center',
          align: 'center',
          color: '#1e293b',
          font: { weight: '700', size: 9, family: 'Inter' },
          backgroundColor: 'rgba(255, 255, 255, 0.88)',
          borderColor: barSeries[0].border,
          borderWidth: 1,
          borderRadius: 4,
          padding: { top: 1, bottom: 1, left: 3, right: 3 },
          formatter: (v) => v > 0 ? `'25: ${formatNumberBadge(v)}` : ''
        }
      },
      {
        type: 'bar',
        label: barSeries[1].label,
        data: barSeries[1].data,
        backgroundColor: barSeries[1].bg,
        borderColor: barSeries[1].border,
        borderWidth: 1.5,
        borderRadius: 6,
        barPercentage: 0.75,
        categoryPercentage: 0.8,
        datalabels: {
          display: showLabels,
          // Año actual: colocado arriba de la barra con insignia destacada
          anchor: 'end',
          align: 'top',
          offset: 3,
          color: barSeries[1].border,
          font: { weight: '800', size: 9.5, family: 'Inter' },
          backgroundColor: 'rgba(255, 255, 255, 0.95)',
          borderColor: barSeries[1].border,
          borderWidth: 1.5,
          borderRadius: 4,
          padding: { top: 1, bottom: 1, left: 4, right: 4 },
          formatter: (v) => v > 0 ? `'26: ${formatNumberBadge(v)}` : ''
        }
      }
    ];

    if (planLine) {
      datasets.push({
        type: 'line',
        label: planLine.label,
        data: planLine.data,
        borderColor: planLine.color,
        borderWidth: 2.5,
        borderDash: [6, 4],
        fill: false,
        pointRadius: 3.5,
        pointBackgroundColor: planLine.color,
        tension: 0.25,
        datalabels: {
          display: showLabels,
          // Plan: colocado arriba de la línea con offset adicional para no tocar las barras
          align: 'top',
          anchor: 'end',
          offset: 6,
          backgroundColor: 'rgba(255, 255, 255, 0.95)',
          borderColor: planLine.color,
          borderWidth: 1.5,
          color: planLine.color,
          font: { weight: '800', size: 9, family: 'Inter' },
          borderRadius: 4,
          padding: { top: 1, bottom: 1, left: 4, right: 4 },
          formatter: (v) => v > 0 ? `Plan: ${formatNumberBadge(v)}` : ''
        }
      });
    }

    charts[canvasId] = new Chart(ctx, {
      data: { labels, datasets },
      options: getBaseChartOptions(showLabels)
    });
  }

  function renderAreaChart(canvasId, labels, areaSeries, planLine, showLabels) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    // areaSeries[0] = 2025 (Año Anterior)
    // areaSeries[1] = 2026 (Año Actual)
    const datasets = [
      {
        type: 'line',
        label: areaSeries[0].label,
        data: areaSeries[0].data,
        backgroundColor: areaSeries[0].bg,
        borderColor: areaSeries[0].border,
        borderWidth: 2.5,
        fill: true,
        tension: 0.35,
        pointRadius: 3.5,
        pointBackgroundColor: areaSeries[0].border,
        datalabels: {
          display: showLabels,
          // Año 2025: etiqueta colocada DEBAJO del punto para CERO colisión con 2026
          anchor: 'start',
          align: 'bottom',
          offset: 6,
          color: '#475569',
          font: { weight: '700', size: 9, family: 'Inter' },
          backgroundColor: 'rgba(255, 255, 255, 0.92)',
          borderColor: areaSeries[0].border,
          borderWidth: 1,
          borderRadius: 4,
          padding: { top: 1, bottom: 1, left: 3, right: 3 },
          formatter: (v) => v > 0 ? `'25: ${formatNumberBadge(v)}` : ''
        }
      },
      {
        type: 'line',
        label: areaSeries[1].label,
        data: areaSeries[1].data,
        backgroundColor: areaSeries[1].bg,
        borderColor: areaSeries[1].border,
        borderWidth: 2.5,
        fill: true,
        tension: 0.35,
        pointRadius: 4,
        pointBackgroundColor: areaSeries[1].border,
        datalabels: {
          display: showLabels,
          // Año 2026: etiqueta colocada ARRIBA del punto con color distintivo
          anchor: 'end',
          align: 'top',
          offset: 6,
          color: areaSeries[1].border,
          font: { weight: '800', size: 9.5, family: 'Inter' },
          backgroundColor: 'rgba(255, 255, 255, 0.95)',
          borderColor: areaSeries[1].border,
          borderWidth: 1.5,
          borderRadius: 4,
          padding: { top: 1, bottom: 1, left: 4, right: 4 },
          formatter: (v) => v > 0 ? `'26: ${formatNumberBadge(v)}` : ''
        }
      }
    ];

    if (planLine) {
      datasets.push({
        type: 'line',
        label: planLine.label,
        data: planLine.data,
        borderColor: planLine.color,
        borderWidth: 2.5,
        borderDash: [6, 4],
        fill: false,
        pointRadius: 3.5,
        pointBackgroundColor: planLine.color,
        tension: 0.25,
        datalabels: {
          display: showLabels,
          // Plan: colocado arriba de la línea con offset y prefijo Plan:
          align: 'top',
          anchor: 'end',
          offset: 10,
          backgroundColor: 'rgba(255, 255, 255, 0.95)',
          borderColor: planLine.color,
          borderWidth: 1.5,
          color: planLine.color,
          font: { weight: '800', size: 9, family: 'Inter' },
          borderRadius: 4,
          padding: { top: 1, bottom: 1, left: 4, right: 4 },
          formatter: (v) => v > 0 ? `Plan: ${formatNumberBadge(v)}` : ''
        }
      });
    }

    charts[canvasId] = new Chart(ctx, {
      data: { labels, datasets },
      options: getBaseChartOptions(showLabels)
    });
  }

  // ══════════════════════════════════════════════
  // TABLE RENDERER WITH YOY BADGES
  // ══════════════════════════════════════════════
  function renderTable(prefix, weekNumbers, currentYear, prevYear, dataMap) {
    const container = document.getElementById(`table${prefix}Container`);
    if (!container) return;

    let html = `
      <div class="table-wrapper-title">
        <i class="fa-solid fa-table-list text-primary"></i>
        <span>Detalle Comparativo Semanal (${currentYear} vs ${prevYear})</span>
      </div>
      <table class="data-table">
        <thead>
          <tr>
            <th style="min-width: 140px;">Métrica / Proceso</th>
    `;
    
    weekNumbers.forEach(w => {
      html += `<th>S${w} (${currentYear})</th>`;
    });
    html += `<th>Total / Prom.</th></tr></thead><tbody>`;

    const metrics = [
      { key: 'recibo', label: '📦 RECIBO', color: '#2563eb' },
      { key: 'despacho', label: '🚛 DESPACHO', color: '#ea580c' },
      { key: 'inventario', label: '📊 INVENTARIO', color: '#059669' }
    ];

    metrics.forEach(m => {
      // Fila Actual
      html += `<tr><td style="color:${m.color}; font-weight:700;">${m.label} ${currentYear}</td>`;
      let sumCurr = 0;
      weekNumbers.forEach(w => {
        const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][m.key] : 0;
        sumCurr += val;
        html += `<td style="font-weight:600;">${Math.round(val).toLocaleString('es-PE')}</td>`;
      });
      const avgOrSumCurr = m.key === 'inventario' ? (sumCurr / weekNumbers.length) : sumCurr;
      html += `<td style="font-weight:800; background:#f8fafc;">${Math.round(avgOrSumCurr).toLocaleString('es-PE')}</td></tr>`;

      // Fila Año Anterior
      html += `<tr style="color:#64748b;"><td style="font-weight:600; padding-left: 20px;">└ Año ${prevYear}</td>`;
      let sumPrev = 0;
      weekNumbers.forEach(w => {
        const val = dataMap[`${prevYear}-${w}`] ? dataMap[`${prevYear}-${w}`][m.key] : 0;
        sumPrev += val;
        html += `<td>${Math.round(val).toLocaleString('es-PE')}</td>`;
      });
      const avgOrSumPrev = m.key === 'inventario' ? (sumPrev / weekNumbers.length) : sumPrev;
      html += `<td style="font-weight:700; background:#f8fafc;">${Math.round(avgOrSumPrev).toLocaleString('es-PE')}</td></tr>`;
    });

    html += '</tbody></table>';
    container.innerHTML = html;
  }
});
