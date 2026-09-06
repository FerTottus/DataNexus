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

  // ══════════════════════════════════════════════
  // PROYECCIÓN A FUTURO (+4 SEMANAS PLAN)
  // ══════════════════════════════════════════════
  const toggleFutureEl = document.getElementById('toggleFutureWeeks');
  const weeksFilterEl  = document.getElementById('weeksFilter');

  toggleFutureEl?.addEventListener('change', (e) => {
    if (!weeksFilterEl) return;
    if (!e.target.checked) {
      // Si desmarca el toggle y estaba en una opción "_plus_4", pasamos a la versión cerrada
      if (weeksFilterEl.value === 'current_plus_4') weeksFilterEl.value = 'current';
      else if (weeksFilterEl.value === '4_plus_4') weeksFilterEl.value = '4';
      else if (weeksFilterEl.value === '8_plus_4') weeksFilterEl.value = '8';
    } else {
      // Si activa el toggle y tenía una opción cerrada simple, pasamos a la opción proyectada
      if (weeksFilterEl.value === 'current') weeksFilterEl.value = 'current_plus_4';
      else if (weeksFilterEl.value === '4') weeksFilterEl.value = '4_plus_4';
      else if (weeksFilterEl.value === '8') weeksFilterEl.value = '8_plus_4';
    }
    renderAll();
  });

  weeksFilterEl?.addEventListener('change', (e) => {
    if (toggleFutureEl) {
      if (e.target.value.includes('plus_4')) {
        toggleFutureEl.checked = true;
      }
    }
    renderAll();
  });

  // ══════════════════════════════════════════════
  // APARTADO DE DIVISIONES: INTERACTIVIDAD
  // ══════════════════════════════════════════════
  // 1. Filtrar por CD (Todas, Secos, Frescos)
  document.querySelectorAll('.div-filter-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('.div-filter-pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      const targetFilter = pill.dataset.divFilter;
      document.querySelectorAll('.division-card').forEach(card => {
        if (targetFilter === 'all' || card.dataset.cd === targetFilter) {
          card.style.display = 'flex';
        } else {
          card.style.display = 'none';
        }
      });
    });
  });

  // 2. Clic en tarjeta de división: activa el tab respectivo y hace scroll suave
  document.querySelectorAll('.division-card').forEach(card => {
    card.addEventListener('click', () => {
      const cd = card.dataset.cd;
      const targetTab = cd === 'frescos' ? 'tab-frescos' : 'tab-secos';
      const targetBtn = document.querySelector(`.glass-tab-btn[data-target="${targetTab}"]`);
      if (targetBtn) targetBtn.click();
      const tabEl = document.getElementById(targetTab);
      if (tabEl) tabEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });

  // 3. Botón de colapso/expansión de la grilla de divisiones
  document.getElementById('btnToggleDivisions')?.addEventListener('click', () => {
    const grid = document.getElementById('divisionsGrid');
    const icon = document.getElementById('iconToggleDivisions');
    if (!grid || !icon) return;
    const isClosed = grid.style.display === 'none';
    grid.style.display = isClosed ? 'grid' : 'none';
    icon.className = isClosed ? 'fa-solid fa-chevron-up' : 'fa-solid fa-chevron-down';
  });

  // Estado del modo KPI: 'periodo' = rango completo, 'semana' = última semana con data
  let kpiMode = 'periodo';
  document.getElementById('btnKpiPeriodo')?.addEventListener('click', () => {
    kpiMode = 'periodo';
    document.getElementById('btnKpiPeriodo').classList.add('kpi-mode-active');
    document.getElementById('btnKpiSemana').classList.remove('kpi-mode-active');
    renderAll();
  });
  document.getElementById('btnKpiSemana')?.addEventListener('click', () => {
    kpiMode = 'semana';
    document.getElementById('btnKpiSemana').classList.add('kpi-mode-active');
    document.getElementById('btnKpiPeriodo').classList.remove('kpi-mode-active');
    renderAll();
  });

  function renderAll() {
    if (!window.dataFrescos || !window.dataSecos) return;
    const numWeeks = document.getElementById('weeksFilter')?.value || 'current_plus_4';
    const showLabels = document.getElementById('toggleDataLabels')?.checked ?? true;
    const includeFuture = document.getElementById('toggleFutureWeeks')?.checked ?? true;

    // CD Frescos
    renderSection('Frescos', window.dataFrescos, resolveColumns(window.dataFrescos.headers, window.dataFrescos.rows), numWeeks, showLabels, includeFuture);

    // CD Secos
    renderSection('Secos', window.dataSecos, resolveColumns(window.dataSecos.headers, window.dataSecos.rows), numWeeks, showLabels, includeFuture);
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

  function renderSection(prefix, data, cols, filterValue, showLabels, includeFuture = true) {
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

    // 4. Semanas a graficar (cerradas + proyección a 4 semanas futuras)
    let isProjectionActive = includeFuture;
    let baseFilter = filterValue;

    if (filterValue === 'current_plus_4') {
      baseFilter = 'current';
      isProjectionActive = true;
    } else if (filterValue === '4_plus_4') {
      baseFilter = '4';
      isProjectionActive = true;
    } else if (filterValue === '8_plus_4') {
      baseFilter = '8';
      isProjectionActive = true;
    }

    let closedWeeks = [];
    if (baseFilter === 'all') {
      closedWeeks = currentYearRows
        .filter(r => r.recibo > 0 || r.despacho > 0 || r.inventario > 0)
        .map(r => r.week);
    } else if (baseFilter === 'current') {
      closedWeeks = [lastDataWeek];
    } else {
      const n = parseInt(baseFilter, 10) || 8;
      const startWeek = Math.max(1, lastDataWeek - n + 1);
      for (let w = startWeek; w <= lastDataWeek; w++) {
        closedWeeks.push(w);
      }
    }

    const weekNumbers = [...closedWeeks];
    if (isProjectionActive) {
      // Agregar exactamente 4 semanas futuras a partir de la semana actual con data
      for (let i = 1; i <= 4; i++) {
        const nextW = lastDataWeek + i;
        if (nextW <= 52 && !weekNumbers.includes(nextW)) {
          weekNumbers.push(nextW);
        }
      }
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

    // Para semanas futuras (w > lastDataWeek), el valor real debe ser null
    // para que Chart.js NO dibuje una barra vacía o en cero y solo se trace el PLAN y el año anterior
    const prevRecibo =       weekNumbers.map(w => get(prevYear, w, 'recibo'));
    const currRecibo =       weekNumbers.map(w => w > lastDataWeek ? null : get(currentYear, w, 'recibo'));
    const planRecibo =       weekNumbers.map(w => get(currentYear, w, 'planRecibo'));

    const prevDespacho =     weekNumbers.map(w => get(prevYear, w, 'despacho'));
    const currDespacho =     weekNumbers.map(w => w > lastDataWeek ? null : get(currentYear, w, 'despacho'));
    const planDespacho =     weekNumbers.map(w => get(currentYear, w, 'planDespacho'));

    const prevInventario =   weekNumbers.map(w => get(prevYear, w, 'inventario'));
    const currInventario =   weekNumbers.map(w => w > lastDataWeek ? null : get(currentYear, w, 'inventario'));
    const planInv =          weekNumbers.map(w => get(currentYear, w, 'planInv'));

    // 6. Renderizar KPI Executive Strip (respetando semanas cerradas para no distorsionar promedios)
    renderKPIs(prefix, weekNumbers, currentYear, prevYear, dataMap, lastDataWeek);

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

    // 📊 INVENTARIO: Grouped bars — barras agrupadas evitan solapamiento de etiquetas
    renderBarChart(`chart${prefix}Inventario`, labels, [
      { label: `Inventario ${prevYear}`, data: prevInventario, bg: 'rgba(148, 163, 184, 0.65)', border: '#94a3b8' },
      { label: `Inventario ${currentYear}`, data: currInventario, bg: '#059669', border: '#047857' },
    ], planInv.some(v => v > 0) ? { label: 'PLAN INV', data: planInv, color: '#0d9488' } : null, showLabels);

    // 8. Tabla Resumen (Muestra hasta 10 semanas para máxima claridad, incluyendo las 4 proyectadas)
    const tableWeeks = weekNumbers.length <= 10 ? weekNumbers : weekNumbers.slice(-10);
    renderTable(prefix, tableWeeks, currentYear, prevYear, dataMap, lastDataWeek);
  }

  // ══════════════════════════════════════════════
  // KPI EXECUTIVE SUMMARY RENDERER
  // ══════════════════════════════════════════════
  function renderKPIs(prefix, weekNumbers, currentYear, prevYear, dataMap, lastDataWeek) {
    const container = document.getElementById(`kpi${prefix}Container`);
    if (!container) return;

    const calcYoY = (curr, prev) => {
      if (prev === 0) return { pct: '0.0%', isUp: true };
      const diff = ((curr - prev) / prev) * 100;
      return { pct: (diff >= 0 ? '+' : '') + diff.toFixed(1) + '%', isUp: diff >= 0 };
    };
    const fmt = n => Math.round(n).toLocaleString('es-PE');

    let reciboCurr, reciboPrev, despachoCurr, despachoPrev, invCurr, invPrev;
    let kpiLabel, semanaRef;

    const closedWeeks = weekNumbers.filter(w => w <= lastDataWeek);
    const futureWeeksCount = weekNumbers.filter(w => w > lastDataWeek).length;

    if (kpiMode === 'semana') {
      // ── MODO SEMANA ACTUAL: solo la última semana con datos reales ───
      const w = lastDataWeek;
      semanaRef = `S${w}`;
      const c = dataMap[`${currentYear}-${w}`] || {};
      const p = dataMap[`${prevYear}-${w}`] || {};
      reciboCurr   = c.recibo || 0;
      reciboPrev   = p.recibo || 0;
      despachoCurr = c.despacho || 0;
      despachoPrev = p.despacho || 0;
      invCurr      = c.inventario || 0;
      invPrev      = p.inventario || 0;
      kpiLabel = `Semana ${w} con datos (${currentYear} vs ${prevYear})${futureWeeksCount > 0 ? ` · +${futureWeeksCount} sem. proyectadas en gráficos` : ''}`;
    } else {
      // ── MODO PERÍODO COMPLETO: suma/promedio de semanas cerradas ─────
      let totReciboCurr = 0, totReciboPrev = 0;
      let totDespachoCurr = 0, totDespachoPrev = 0;
      let totInvCurr = 0, totInvPrev = 0, countInv = 0;

      closedWeeks.forEach(w => {
        const c = dataMap[`${currentYear}-${w}`];
        const p = dataMap[`${prevYear}-${w}`];
        if (c) {
          totReciboCurr   += c.recibo || 0;
          totDespachoCurr += c.despacho || 0;
          if (c.inventario > 0) { totInvCurr += c.inventario; countInv++; }
        }
        if (p) {
          totReciboPrev   += p.recibo || 0;
          totDespachoPrev += p.despacho || 0;
          totInvPrev      += p.inventario || 0;
        }
      });

      reciboCurr   = totReciboCurr;
      reciboPrev   = totReciboPrev;
      despachoCurr = totDespachoCurr;
      despachoPrev = totDespachoPrev;
      invCurr      = countInv > 0 ? totInvCurr / countInv : 0;
      invPrev      = closedWeeks.length > 0 ? totInvPrev / closedWeeks.length : 0;
      const proyNote = futureWeeksCount > 0 ? ` + ${futureWeeksCount} proyectadas (Plan)` : '';
      kpiLabel = `Últimas ${closedWeeks.length} semanas cerradas (${currentYear} vs ${prevYear})${proyNote}`;
      semanaRef = null;
    }

    const yoyRecibo   = calcYoY(reciboCurr, reciboPrev);
    const yoyDespacho = calcYoY(despachoCurr, despachoPrev);
    const yoyInv      = calcYoY(invCurr, invPrev);

    const modeIsWeek = kpiMode === 'semana';
    const invLabel = modeIsWeek ? 'Inventario Final' : 'Stock Inventario (Prom.)';
    const reciboLabel = modeIsWeek ? `Recibo S${lastDataWeek}` : 'Total Entradas (Recibo)';
    const despachoLabel = modeIsWeek ? `Despacho S${lastDataWeek}` : 'Total Salidas (Despacho)';

    container.innerHTML = `
      <div style="grid-column:1/-1; font-size:0.78rem; font-weight:700; color:#64748b; margin-bottom:-6px; padding-left:2px; display:flex; align-items:center; gap:8px;">
        <i class="fa-solid fa-chart-pie" style="color:#2563eb;"></i> ${kpiLabel}
      </div>

      <div class="kpi-card kpi-recibo">
        <div class="kpi-title"><i class="fa-solid fa-boxes-packing text-primary"></i> ${reciboLabel}</div>
        <div class="kpi-value">${fmt(reciboCurr)}</div>
        <div class="kpi-sub">
          <span class="kpi-badge ${yoyRecibo.isUp ? 'badge-up' : 'badge-down'}">
            <i class="fa-solid fa-arrow-${yoyRecibo.isUp ? 'trend-up' : 'trend-down'}"></i> ${yoyRecibo.pct}
          </span>
          <span>vs ${fmt(reciboPrev)} (${prevYear})</span>
        </div>
      </div>

      <div class="kpi-card kpi-despacho">
        <div class="kpi-title"><i class="fa-solid fa-truck-fast" style="color:#ea580c;"></i> ${despachoLabel}</div>
        <div class="kpi-value">${fmt(despachoCurr)}</div>
        <div class="kpi-sub">
          <span class="kpi-badge ${yoyDespacho.isUp ? 'badge-up' : 'badge-down'}">
            <i class="fa-solid fa-arrow-${yoyDespacho.isUp ? 'trend-up' : 'trend-down'}"></i> ${yoyDespacho.pct}
          </span>
          <span>vs ${fmt(despachoPrev)} (${prevYear})</span>
        </div>
      </div>

      <div class="kpi-card kpi-inventario">
        <div class="kpi-title"><i class="fa-solid fa-warehouse" style="color:#059669;"></i> ${invLabel}</div>
        <div class="kpi-value">${fmt(invCurr)}</div>
        <div class="kpi-sub">
          <span class="kpi-badge ${yoyInv.isUp ? 'badge-up' : 'badge-down'}">
            <i class="fa-solid fa-arrow-${yoyInv.isUp ? 'trend-up' : 'trend-down'}"></i> ${yoyInv.pct}
          </span>
          <span>vs ${fmt(invPrev)} (${prevYear})</span>
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

  // ── Fuentes responsivas para datalabels ──────────────────────────────────
  // Escala con el ancho real del canvas. Valores mínimos y máximos garantizan
  // legibilidad tanto en móvil (350px) como en pantallas 4K (2400px+).
  //   divisor bajo  → letra más grande en relación al ancho del chart
  //   clamp min/max → nunca demasiado pequeño ni demasiado enorme
  const dlFont = {
    // Barra año anterior (texto oscuro sobre barra clara)
    barPrev: (ctx) => ({
      weight: '700',
      size: Math.max(10, Math.min(16, ctx.chart.width / 62)),
      family: 'Inter'
    }),
    // Barra año actual (texto blanco sobre barra sólida)
    barCurr: (ctx) => ({
      weight: '800',
      size: Math.max(11, Math.min(17, ctx.chart.width / 56)),
      family: 'Inter'
    }),
    // Línea Plan (pill encima de la línea punteada)
    plan: (ctx) => ({
      weight: '700',
      size: Math.max(10, Math.min(14, ctx.chart.width / 72)),
      family: 'Inter'
    }),
    // Área / línea año anterior
    areaPrev: (ctx) => ({
      weight: '700',
      size: Math.max(10, Math.min(15, ctx.chart.width / 65)),
      family: 'Inter'
    }),
    // Área / línea año actual
    areaCurr: (ctx) => ({
      weight: '800',
      size: Math.max(11, Math.min(16, ctx.chart.width / 58)),
      family: 'Inter'
    })
  };

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
              if (ctx.parsed.y === null || ctx.parsed.y === undefined) {
                return `  ${label}: -- (Pendiente / Proyectado)`;
              }
              const val = Math.round(ctx.parsed.y).toLocaleString('es-PE');
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

    // barSeries[0] = Año Anterior (barra clara)
    // barSeries[1] = Año Actual (barra sólida)
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
          // Dentro de la barra, en el centro — texto oscuro sobre barra clara
          anchor: 'center',
          align: 'center',
          color: '#1e293b',
          font: dlFont.barPrev,
          backgroundColor: null,
          borderWidth: 0,
          formatter: (v) => (v !== null && v > 0) ? formatNumberBadge(v) : ''
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
          // Dentro de la barra sólida, centrado — texto blanco siempre legible
          anchor: 'center',
          align: 'center',
          color: '#ffffff',
          font: dlFont.barCurr,
          backgroundColor: null,
          borderWidth: 0,
          formatter: (v) => (v !== null && v > 0) ? formatNumberBadge(v) : ''
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
          // Plan encima de la línea punteada
          align: 'top',
          anchor: 'end',
          offset: 4,
          backgroundColor: 'rgba(255,255,255,0.92)',
          borderColor: planLine.color,
          borderWidth: 1,
          color: planLine.color,
          font: dlFont.plan,
          borderRadius: 3,
          padding: { top: 2, bottom: 2, left: 5, right: 5 },
          formatter: (v) => (v !== null && v > 0) ? `P: ${formatNumberBadge(v)}` : ''
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
          // Año anterior: etiqueta debajo del punto
          anchor: 'start',
          align: 'bottom',
          offset: 6,
          color: '#475569',
          font: dlFont.areaPrev,
          backgroundColor: 'rgba(255, 255, 255, 0.92)',
          borderColor: areaSeries[0].border,
          borderWidth: 1,
          borderRadius: 4,
          padding: { top: 2, bottom: 2, left: 5, right: 5 },
          formatter: (v) => v > 0 ? formatNumberBadge(v) : ''
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
          // Año actual: etiqueta arriba del punto con color de la serie
          anchor: 'end',
          align: 'top',
          offset: 6,
          color: areaSeries[1].border,
          font: dlFont.areaCurr,
          backgroundColor: 'rgba(255, 255, 255, 0.9)',
          borderColor: areaSeries[1].border,
          borderWidth: 1.5,
          borderRadius: 4,
          padding: { top: 2, bottom: 2, left: 5, right: 5 },
          formatter: (v) => v > 0 ? formatNumberBadge(v) : ''
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
          align: 'top',
          anchor: 'end',
          offset: 10,
          backgroundColor: 'rgba(255,255,255,0.92)',
          borderColor: planLine.color,
          borderWidth: 1,
          color: planLine.color,
          font: dlFont.plan,
          borderRadius: 3,
          padding: { top: 2, bottom: 2, left: 5, right: 5 },
          formatter: (v) => v > 0 ? `P: ${formatNumberBadge(v)}` : ''
        }
      });
    }

    charts[canvasId] = new Chart(ctx, {
      data: { labels, datasets },
      options: getBaseChartOptions(showLabels)
    });
  }

  // ══════════════════════════════════════════════
  // TABLE RENDERER WITH YOY BADGES & PLAN PROJECTION
  // ══════════════════════════════════════════════
  function renderTable(prefix, weekNumbers, currentYear, prevYear, dataMap, lastDataWeek) {
    const container = document.getElementById(`table${prefix}Container`);
    if (!container) return;

    let html = `
      <div class="table-wrapper-title">
        <i class="fa-solid fa-table-list text-primary"></i>
        <span>Detalle Comparativo Semanal (${currentYear} vs ${prevYear}) + Proyección Plan</span>
      </div>
      <table class="data-table">
        <thead>
          <tr>
            <th style="min-width: 150px;">Métrica / Proceso</th>
    `;
    
    weekNumbers.forEach(w => {
      const isFuture = w > lastDataWeek;
      if (isFuture) {
        html += `<th style="background:#eff6ff; color:#1d4ed8; border-bottom: 2px solid #3b82f6;">S${w} <span style="font-size:0.72rem; font-weight:700; color:#2563eb; display:block;">(Plan Proy)</span></th>`;
      } else {
        html += `<th>S${w} (${currentYear})</th>`;
      }
    });
    html += `<th style="min-width:130px; background:#f8fafc;">Total / Prom. (Cerradas)</th></tr></thead><tbody>`;

    const metrics = [
      { key: 'recibo', planKey: 'planRecibo', label: '📦 RECIBO', color: '#2563eb' },
      { key: 'despacho', planKey: 'planDespacho', label: '🚛 DESPACHO', color: '#ea580c' },
      { key: 'inventario', planKey: 'planInv', label: '📊 INVENTARIO', color: '#059669' }
    ];

    metrics.forEach(m => {
      // 1. Fila Actual (Real)
      html += `<tr><td style="color:${m.color}; font-weight:700;">${m.label} ${currentYear} (Real)</td>`;
      let sumCurr = 0;
      let countCurr = 0;
      weekNumbers.forEach(w => {
        const isFuture = w > lastDataWeek;
        if (isFuture) {
          html += `<td style="color:#94a3b8; font-style:italic; background:#f8fafc;">--</td>`;
        } else {
          const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][m.key] : 0;
          sumCurr += val;
          countCurr++;
          html += `<td style="font-weight:600;">${Math.round(val).toLocaleString('es-PE')}</td>`;
        }
      });
      const avgOrSumCurr = m.key === 'inventario' ? (countCurr > 0 ? sumCurr / countCurr : 0) : sumCurr;
      html += `<td style="font-weight:800; background:#f1f5f9;">${Math.round(avgOrSumCurr).toLocaleString('es-PE')}</td></tr>`;

      // 2. Fila Plan (Metas / Proyección)
      html += `<tr style="color:#0284c7; background:rgba(239, 246, 255, 0.35);"><td style="font-weight:600; padding-left: 20px;">└ Plan Proyectado</td>`;
      let sumPlan = 0;
      let countPlan = 0;
      weekNumbers.forEach(w => {
        const isFuture = w > lastDataWeek;
        const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][m.planKey] : 0;
        if (!isFuture) { sumPlan += val; countPlan++; }
        const highlightStyle = isFuture ? 'font-weight:700; color:#1d4ed8; background:rgba(219, 234, 254, 0.5);' : '';
        html += `<td style="${highlightStyle}">${val > 0 ? Math.round(val).toLocaleString('es-PE') : '--'}</td>`;
      });
      const avgOrSumPlan = m.key === 'inventario' ? (countPlan > 0 ? sumPlan / countPlan : 0) : sumPlan;
      html += `<td style="font-weight:700; background:#eff6ff;">${Math.round(avgOrSumPlan).toLocaleString('es-PE')}</td></tr>`;

      // 3. Fila Año Anterior (Real)
      html += `<tr style="color:#64748b;"><td style="font-weight:600; padding-left: 20px;">└ Real ${prevYear}</td>`;
      let sumPrev = 0;
      let countPrev = 0;
      weekNumbers.forEach(w => {
        const isFuture = w > lastDataWeek;
        const val = dataMap[`${prevYear}-${w}`] ? dataMap[`${prevYear}-${w}`][m.key] : 0;
        if (!isFuture) { sumPrev += val; countPrev++; }
        html += `<td>${val > 0 ? Math.round(val).toLocaleString('es-PE') : '--'}</td>`;
      });
      const avgOrSumPrev = m.key === 'inventario' ? (countPrev > 0 ? sumPrev / countPrev : 0) : sumPrev;
      html += `<td style="font-weight:700; background:#f8fafc;">${Math.round(avgOrSumPrev).toLocaleString('es-PE')}</td></tr>`;
    });

    html += '</tbody></table>';
    container.innerHTML = html;
  }
});
