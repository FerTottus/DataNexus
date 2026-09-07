document.addEventListener('DOMContentLoaded', () => {
  // ══════════════════════════════════════════════
  // CHART DATA LABELS PLUGIN REGISTRATION
  // ══════════════════════════════════════════════
  if (window.ChartDataLabels) {
    Chart.register(window.ChartDataLabels);
  }

  // ══════════════════════════════════════════════
  // CHART REGISTRY & RESPONSIVE RESIZE HANDLER
  // ══════════════════════════════════════════════
  const charts = {};

  function resizeAllCharts() {
    Object.values(charts).forEach(chart => {
      if (chart && typeof chart.resize === 'function') {
        chart.resize();
      }
    });
  }

  let resizeTimeout;
  window.addEventListener('resize', () => {
    clearTimeout(resizeTimeout);
    resizeTimeout = setTimeout(resizeAllCharts, 120);
  });

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
      requestAnimationFrame(() => {
        resizeAllCharts();
      });
    });
  });
  document.getElementById('btnChangeSheet')?.addEventListener('click', () => {
    // Resetear input y datos cargados
    const inputEl = document.getElementById('sheetUrlInput');
    if (inputEl) inputEl.value = '';

    window.dataSecos = null;
    window.dataFrescos = null;
    window.parsedDivisionData = null;
    window.dataDivision = null;

    // Destruir instancias de gráficos para liberar memoria
    Object.keys(charts).forEach(key => {
      if (charts[key] && typeof charts[key].destroy === 'function') {
        charts[key].destroy();
      }
      delete charts[key];
    });

    document.getElementById('dashboardSection')?.classList.add('hidden');
    document.getElementById('connectionSuccessInfo')?.classList.add('hidden');
    document.getElementById('connectBox')?.classList.remove('hidden');
  });

  // Permitir presionar Enter en el input para cargar
  document.getElementById('sheetUrlInput')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      document.getElementById('btnFetchData')?.click();
    }
  });

  // Botón para cargar datos de prueba a demanda (opcional para pruebas/demos)
  document.getElementById('btnLoadDemo')?.addEventListener('click', () => {
    window.dataSecos = JSON.parse(JSON.stringify(DEFAULT_SECOS_DATA));
    window.dataFrescos = JSON.parse(JSON.stringify(DEFAULT_FRESCOS_DATA));
    window.parsedDivisionData = JSON.parse(JSON.stringify(DEFAULT_DIVISIONS_DATA));

    document.getElementById('connectBox')?.classList.add('hidden');
    document.getElementById('connectionSuccessInfo')?.classList.remove('hidden');
    const nameEl = document.getElementById('connectedSheetName');
    if (nameEl) nameEl.textContent = 'Datos de prueba (Demostración)';
    document.getElementById('dashboardSection')?.classList.remove('hidden');

    renderAll();
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

      // Cargar desglose de divisiones si la pestaña DIVISION existe
      const sheetDivisionTab = tabNames.find(t => {
        const c = t.trim().toUpperCase().replace('Ó','O');
        return c === 'DIVISION' || c === 'DIVISIONES';
      });
      if (sheetDivisionTab) {
        try {
          window.dataDivision = await window.GoogleSheetsService.fetchSheetData(sheetId, `'${sheetDivisionTab}'!A1:AZ55`);
          window.parsedDivisionData = parseDivisionSheet(window.dataDivision?.rawValues);
        } catch (divErr) {
          console.warn("Aviso: No se pudo cargar hoja DIVISION, usando ratios calculados:", divErr);
        }
      }

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
  // CATÁLOGO Y CONFIGURACIÓN DE DIVISIONES
  // ══════════════════════════════════════════════
  const DIVISION_NAMES = {
    J01: 'PGC COMESTIBLE',
    J02: 'PGC NO COMESTIBLE',
    J03: 'CARNES Y PESCADOS',
    J04: 'FRUTAS Y VERDURAS',
    J05: 'FLC',
    J06: 'PANADERIA Y PASTELERIA',
    J07: 'PLATOS PREPARADOS',
    J08: 'VESTUARIO',
    J09: 'HOGAR',
    J10: 'BAZAR',
    J11: 'ELECTROHOGAR',
    J12: 'INSTITUCIONALES'
  };

  // Centro de Distribución principal asociado a cada división
  const DIVISION_PRIMARY_CD = {
    J01: 'secos',
    J02: 'secos',
    J03: 'frescos',
    J04: 'frescos',
    J05: 'frescos',
    J06: 'frescos',
    J07: 'frescos',
    J08: 'secos',
    J09: 'secos',
    J10: 'secos',
    J11: 'secos',
    J12: 'secos'
  };

  // Participación volumétrica promedio de referencia por CD
  const DIVISION_SHARE = {
    Secos: {
      J01: 0.594, J02: 0.196, J08: 0.093, J09: 0.043,
      J10: 0.025, J05: 0.025, J11: 0.018, J06: 0.004,
      J07: 0.003, J12: 0.0003, J03: 0.0, J04: 0.0
    },
    Frescos: {
      J05: 0.336, J04: 0.254, J06: 0.228, J07: 0.132,
      J03: 0.039, J01: 0.009, J02: 0.0, J08: 0.0,
      J09: 0.0, J10: 0.0, J11: 0.0, J12: 0.0
    }
  };

  // Multi-selección activa de divisiones (código Set)
  const selectedDivisions = new Set();

  // Parser de la hoja DIVISION cuando está presente en Google Sheets
  function parseDivisionSheet(rawValues) {
    if (!rawValues || rawValues.length < 25) return null;
    const divData = { Secos: {}, Frescos: {} };

    // Secos (filas 18 a 30 aprox)
    const headerRowSecos = rawValues[18] || [];
    const secosRecWeeks = [], secosDespWeeks = [], secosInvWeeks = [];
    for (let c = 2; c <= 9; c++) {
      const p = parseWeekLabel(headerRowSecos[c]);
      if (p) secosRecWeeks.push({ col: c, key: `${p.year}-${p.week}` });
    }
    for (let c = 11; c <= 18; c++) {
      const p = parseWeekLabel(headerRowSecos[c]);
      if (p) secosDespWeeks.push({ col: c, key: `${p.year}-${p.week}` });
    }
    for (let c = 20; c <= 27; c++) {
      const p = parseWeekLabel(headerRowSecos[c]);
      if (p) secosInvWeeks.push({ col: c, key: `${p.year}-${p.week}` });
    }

    for (let r = 19; r <= Math.min(32, rawValues.length - 1); r++) {
      const row = rawValues[r] || [];
      const codeRec = String(row[1] || '').trim().toUpperCase();
      if (/^J\d{2}$/.test(codeRec)) {
        if (!divData.Secos[codeRec]) divData.Secos[codeRec] = { recibo: {}, despacho: {}, inventario: {} };
        secosRecWeeks.forEach(w => {
          divData.Secos[codeRec].recibo[w.key] = parseFloat(row[w.col]) || 0;
        });
      }
      const codeDesp = String(row[10] || '').trim().toUpperCase();
      if (/^J\d{2}$/.test(codeDesp)) {
        if (!divData.Secos[codeDesp]) divData.Secos[codeDesp] = { recibo: {}, despacho: {}, inventario: {} };
        secosDespWeeks.forEach(w => {
          divData.Secos[codeDesp].despacho[w.key] = parseFloat(row[w.col]) || 0;
        });
      }
      const codeInv = String(row[19] || '').trim().toUpperCase();
      if (/^J\d{2}$/.test(codeInv)) {
        if (!divData.Secos[codeInv]) divData.Secos[codeInv] = { recibo: {}, despacho: {}, inventario: {} };
        secosInvWeeks.forEach(w => {
          divData.Secos[codeInv].inventario[w.key] = parseFloat(row[w.col]) || 0;
        });
      }
    }

    // Frescos (filas 37 a 46 aprox)
    if (rawValues.length >= 40) {
      const headerRowFrescos = rawValues[37] || [];
      const fRecWeeks = [], fDespWeeks = [], fInvWeeks = [];
      for (let c = 2; c <= 9; c++) {
        const p = parseWeekLabel(headerRowFrescos[c]);
        if (p) fRecWeeks.push({ col: c, key: `${p.year}-${p.week}` });
      }
      for (let c = 11; c <= 18; c++) {
        const p = parseWeekLabel(headerRowFrescos[c]);
        if (p) fDespWeeks.push({ col: c, key: `${p.year}-${p.week}` });
      }
      for (let c = 20; c <= 27; c++) {
        const p = parseWeekLabel(headerRowFrescos[c]);
        if (p) fInvWeeks.push({ col: c, key: `${p.year}-${p.week}` });
      }

      for (let r = 38; r <= Math.min(46, rawValues.length - 1); r++) {
        const row = rawValues[r] || [];
        const codeRec = String(row[1] || '').trim().toUpperCase();
        if (/^J\d{2}$/.test(codeRec)) {
          if (!divData.Frescos[codeRec]) divData.Frescos[codeRec] = { recibo: {}, despacho: {}, inventario: {} };
          fRecWeeks.forEach(w => {
            divData.Frescos[codeRec].recibo[w.key] = parseFloat(row[w.col]) || 0;
          });
        }
        const codeDesp = String(row[10] || '').trim().toUpperCase();
        if (/^J\d{2}$/.test(codeDesp)) {
          if (!divData.Frescos[codeDesp]) divData.Frescos[codeDesp] = { recibo: {}, despacho: {}, inventario: {} };
          fDespWeeks.forEach(w => {
            divData.Frescos[codeDesp].despacho[w.key] = parseFloat(row[w.col]) || 0;
          });
        }
        const codeInv = String(row[19] || '').trim().toUpperCase();
        if (/^J\d{2}$/.test(codeInv)) {
          if (!divData.Frescos[codeInv]) divData.Frescos[codeInv] = { recibo: {}, despacho: {}, inventario: {} };
          fInvWeeks.forEach(w => {
            divData.Frescos[codeInv].inventario[w.key] = parseFloat(row[w.col]) || 0;
          });
        }
      }
    }
    return divData;
  }

  // ══════════════════════════════════════════════
  // ══════════════════════════════════════════════
  // APARTADO DE DIVISIONES: FILTRADO MULTI-SELECCIÓN & INTERACTIVIDAD
  // ══════════════════════════════════════════════
  function updateDivisionUI() {
    const count = selectedDivisions.size;
    const isFiltering = count > 0;
    const isAll = count === 12;

    // 1. Actualizar estado visual de las 12 tarjetas de división
    document.querySelectorAll('.division-card').forEach(card => {
      const cardCode = card.dataset.code;
      if (isFiltering && !isAll) {
        if (selectedDivisions.has(cardCode)) {
          card.classList.add('active-selected');
          card.classList.remove('dimmed');
        } else {
          card.classList.remove('active-selected');
          card.classList.add('dimmed');
        }
      } else if (isAll) {
        card.classList.add('active-selected');
        card.classList.remove('dimmed');
      } else {
        card.classList.remove('active-selected', 'dimmed');
      }
    });

    // 2. Actualizar indicador y badge en el encabezado
    const indicator = document.getElementById('activeDivisionIndicator');
    const nameEl = document.getElementById('activeDivName');
    const badgeEl = document.getElementById('divTotalBadge');
    const btnClear = document.getElementById('btnClearDivisionFilter');

    if (badgeEl) {
      if (isFiltering && !isAll) {
        badgeEl.textContent = `${count} de 12 Seleccionadas`;
        badgeEl.style.background = '#2563eb';
        badgeEl.style.color = '#ffffff';
      } else if (isAll) {
        badgeEl.textContent = '12 de 12 Seleccionadas (Total)';
        badgeEl.style.background = '#059669';
        badgeEl.style.color = '#ffffff';
      } else {
        badgeEl.textContent = '12 Divisiones (4 Filas × 3 Columnas)';
        badgeEl.style.background = '#e0e7ff';
        badgeEl.style.color = '#3730a3';
      }
    }

    if (indicator && nameEl) {
      if (isFiltering) {
        indicator.style.display = 'inline-flex';
        const sortedCodes = Array.from(selectedDivisions).sort();
        if (isAll) {
          nameEl.textContent = 'Todas las divisiones (12)';
        } else if (sortedCodes.length === 1) {
          const c = sortedCodes[0];
          nameEl.textContent = `${c} · ${DIVISION_NAMES[c] || ''}`;
        } else {
          nameEl.textContent = `${sortedCodes.length} divisiones (${sortedCodes.join(', ')})`;
        }
        if (btnClear) {
          btnClear.innerHTML = `<i class="fa-solid fa-xmark"></i> Quitar filtros (${sortedCodes.length})`;
        }
      } else {
        indicator.style.display = 'none';
      }
    }
  }

  window.toggleDivisionFilter = function(code) {
    if (!code) {
      selectedDivisions.clear();
    } else {
      const wasEmpty = selectedDivisions.size === 0;
      if (selectedDivisions.has(code)) {
        selectedDivisions.delete(code);
      } else {
        selectedDivisions.add(code);
        // Si es la primera división que se selecciona, cambiar suavemente al tab de su CD principal
        if (wasEmpty) {
          const primaryCd = DIVISION_PRIMARY_CD[code] || 'secos';
          const targetTab = primaryCd === 'frescos' ? 'tab-frescos' : 'tab-secos';
          const targetBtn = document.querySelector(`.glass-tab-btn[data-target="${targetTab}"]`);
          if (targetBtn && !targetBtn.classList.contains('active')) {
            targetBtn.click();
          }
        }
      }
    }
    updateDivisionUI();
    renderAll();
  };

  window.clearDivisionFilter = function() {
    selectedDivisions.clear();
    updateDivisionUI();
    renderAll();
  };

  // Clic en tarjeta de división: activa/desactiva multi-selección SIN saltos ni redirecciones
  document.querySelectorAll('.division-card').forEach(card => {
    card.addEventListener('click', () => {
      const code = card.dataset.code;
      window.toggleDivisionFilter(code);
    });
  });

  // Botón para limpiar filtro desde el encabezado
  document.getElementById('btnClearDivisionFilter')?.addEventListener('click', (e) => {
    e.stopPropagation();
    window.clearDivisionFilter();
  });

  // Botón para seleccionar todas las divisiones corporativas
  document.getElementById('btnSelectAllDivisions')?.addEventListener('click', () => {
    Object.keys(DIVISION_NAMES).forEach(c => selectedDivisions.add(c));
    updateDivisionUI();
    renderAll();
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
  function parseWeekLabel(label) {
    const s = String(label).replace(/[\[\]]/g, '');
    const parts = s.split('-');
    if (parts.length !== 2) return null;
    return { week: parseInt(parts[0], 10), year: parseInt(parts[1], 10) };
  }

  function renderSection(prefix, data, cols, filterValue, showLabels, includeFuture = true) {
    if (!data.rows || data.rows.length === 0) return;

    // Determinar si hay filtro de división activo
    const countDivs = selectedDivisions.size;
    const isFilteringDivisions = countDivs > 0 && countDivs < 12;

    // Banner de división en este tab
    const bannerEl = document.getElementById(`divFilterBanner${prefix}`);
    if (bannerEl) {
      if (isFilteringDivisions) {
        bannerEl.style.display = 'flex';
        const sortedCodes = Array.from(selectedDivisions).sort();
        const pillsHtml = sortedCodes.map(c => `
          <span class="filter-pill-tag">
            <strong>${c}</strong> <span style="opacity:0.9;">${DIVISION_NAMES[c] || ''}</span>
            <i class="fa-solid fa-xmark remove-div-pill" data-code="${c}" title="Deseleccionar ${c}"></i>
          </span>
        `).join('');

        bannerEl.innerHTML = `
          <div class="banner-text" style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
            <div style="display:flex; align-items:center; gap:6px; font-weight:700;">
              <i class="fa-solid fa-filter text-primary"></i>
              <span>Filtrando ${countDivs} ${countDivs === 1 ? 'división' : 'divisiones'} en CD ${prefix.toUpperCase()}:</span>
            </div>
            <div style="display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
              ${pillsHtml}
            </div>
          </div>
          <button class="btn-clear-banner" onclick="window.clearDivisionFilter()">
            <i class="fa-solid fa-xmark"></i> Quitar filtros (${countDivs})
          </button>
        `;

        bannerEl.querySelectorAll('.remove-div-pill').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            window.toggleDivisionFilter(btn.dataset.code);
          });
        });
      } else {
        bannerEl.style.display = 'none';
        bannerEl.innerHTML = '';
      }
    }

    // Participación acumulada de las divisiones seleccionadas en este CD (para plan y fallback)
    let combinedShare = 1;
    if (isFilteringDivisions) {
      combinedShare = 0;
      selectedDivisions.forEach(code => {
        combinedShare += (DIVISION_SHARE[prefix]?.[code] || 0);
      });
    }

    // 1. Extraer y estructurar datos
    const allParsed = [];
    data.rows.forEach(row => {
      const label = row[cols.timeCol];
      if (!label) return;
      const parsed = parseWeekLabel(label);
      if (!parsed || isNaN(parsed.week) || isNaN(parsed.year)) return;

      const weekKey = `${parsed.year}-${parsed.week}`;
      const baseRec      = parseFloat(row[cols.recibo]) || 0;
      const baseDesp     = parseFloat(row[cols.despacho]) || 0;
      const baseInv      = parseFloat(row[cols.inventario]) || 0;
      const basePlanRec  = parseFloat(row[cols.planRecibo]) || 0;
      const basePlanDesp = parseFloat(row[cols.planDespacho]) || 0;
      const basePlanInv  = parseFloat(row[cols.planInv]) || 0;

      let rRec  = baseRec;
      let rDesp = baseDesp;
      let rInv  = baseInv;
      let pRec  = basePlanRec;
      let pDesp = basePlanDesp;
      let pInv  = basePlanInv;

      if (isFilteringDivisions) {
        let sumRec = 0;
        let sumDesp = 0;
        let sumInv = 0;

        selectedDivisions.forEach(code => {
          const divExact = window.parsedDivisionData?.[prefix]?.[code];
          const divShare = DIVISION_SHARE[prefix]?.[code] || 0;

          // Recibo
          if (divExact?.recibo && divExact.recibo[weekKey] !== undefined) {
            sumRec += divExact.recibo[weekKey];
          } else {
            sumRec += baseRec * divShare;
          }

          // Despacho
          if (divExact?.despacho && divExact.despacho[weekKey] !== undefined) {
            sumDesp += divExact.despacho[weekKey];
          } else {
            sumDesp += baseDesp * divShare;
          }

          // Inventario
          if (divExact?.inventario && divExact.inventario[weekKey] !== undefined) {
            sumInv += divExact.inventario[weekKey];
          } else {
            sumInv += baseInv * divShare;
          }
        });

        rRec  = sumRec;
        rDesp = sumDesp;
        rInv  = sumInv;
        pRec  = basePlanRec * combinedShare;
        pDesp = basePlanDesp * combinedShare;
        pInv  = basePlanInv * combinedShare;
      }

      allParsed.push({
        label,
        week: parsed.week,
        year: parsed.year,
        recibo:       rRec,
        despacho:     rDesp,
        inventario:   rInv,
        planRecibo:   pRec,
        planDespacho: pDesp,
        planInv:      pInv
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

    const labels = weekNumbers.map(w => {
      const mIdx = getMonthForWeek(w, currentYear);
      const mShort = MONTH_NAMES_SHORT[mIdx] || '';
      return [`S${w}`, mShort];
    });
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
    ], planRecibo.some(v => v > 0) ? { label: 'PLAN RECIBO', data: planRecibo, color: '#0284c7' } : null, showLabels, currentYear);

    // 🚛 DESPACHO: Coral / Flame Orange Palette + Ruby Plan
    renderBarChart(`chart${prefix}Despacho`, labels, [
      { label: `Despacho ${prevYear}`, data: prevDespacho, bg: 'rgba(253, 186, 116, 0.75)', border: '#fb923c' },
      { label: `Despacho ${currentYear}`, data: currDespacho, bg: '#ea580c', border: '#c2410c' },
    ], planDespacho.some(v => v > 0) ? { label: 'PLAN DESPACHO', data: planDespacho, color: '#e11d48' } : null, showLabels, currentYear);

    // 📊 INVENTARIO: Grouped bars — barras agrupadas con espacio suficiente
    renderBarChart(`chart${prefix}Inventario`, labels, [
      { label: `Inventario ${prevYear}`, data: prevInventario, bg: 'rgba(148, 163, 184, 0.65)', border: '#94a3b8' },
      { label: `Inventario ${currentYear}`, data: currInventario, bg: '#059669', border: '#047857' },
    ], planInv.some(v => v > 0) ? { label: 'PLAN INV', data: planInv, color: '#0d9488' } : null, showLabels, currentYear);

    // 8. Tabla Resumen: El usuario solicitó explícitamente NO incluir la proyección futura en la tabla,
    // sino mostrar únicamente las semanas cerradas reales (específicamente las 8 semanas cerradas).
    const tableWeeks = closedWeeks.length <= 8 ? closedWeeks : closedWeeks.slice(-8);
    renderTable(prefix, tableWeeks, currentYear, prevYear, dataMap, lastDataWeek);

    // 9. Tabla de Movimientos por División (Año Actual) con Flechas de Tendencia
    renderDivisionsTable(prefix, tableWeeks, currentYear, dataMap, lastDataWeek);
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

    let divTag = '';
    if (selectedDivisions.size > 0 && selectedDivisions.size < 12) {
      const sorted = Array.from(selectedDivisions).sort();
      if (sorted.length === 1) {
        divTag = ` [División: ${sorted[0]} · ${DIVISION_NAMES[sorted[0]] || ''}]`;
      } else {
        divTag = ` [${sorted.length} Divisiones: ${sorted.join(', ')}]`;
      }
    }

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
      kpiLabel = `Semana ${w} con datos (${currentYear} vs ${prevYear})${divTag}${futureWeeksCount > 0 ? ` · +${futureWeeksCount} sem. proyectadas en gráficos` : ''}`;
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
      kpiLabel = `Últimas ${closedWeeks.length} semanas cerradas (${currentYear} vs ${prevYear})${divTag}${proyNote}`;
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
  // HELPER CALENDAR & MONTH MAPPING (ISO-8601)
  // ══════════════════════════════════════════════
  const MONTH_NAMES_SHORT = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Set', 'Oct', 'Nov', 'Dic'];
  const MONTH_NAMES_FULL = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre'];

  function getMonthForWeek(week, year = 2026) {
    const y = parseInt(year, 10) || 2026;
    const w = parseInt(week, 10) || 1;
    // Anclaje en el jueves de la primera semana ISO del año (criterio ISO-8601 estándar)
    const jan4 = new Date(y, 0, 4);
    const dayOfWeek = jan4.getDay() || 7;
    const thursdayOfW1 = new Date(y, 0, 4 + (4 - dayOfWeek));
    const targetDate = new Date(thursdayOfW1.getTime() + (w - 1) * 7 * 86400000);
    return targetDate.getMonth();
  }

  // ══════════════════════════════════════════════
  // CHART BUILDERS WITH DATALABELS & LIQUID THEME
  // ══════════════════════════════════════════════
  function formatNumberBadge(val) {
    if (!val || val === 0) return '';
    if (val >= 1000000) {
      const m = val / 1000000;
      return (m >= 10 ? m.toFixed(1) : m.toFixed(2)) + 'M';
    }
    if (val >= 1000) {
      return Math.round(val / 1000).toLocaleString('es-PE') + 'k';
    }
    return Math.round(val).toLocaleString('es-PE');
  }

  // ── Fuentes responsivas calibradas para que NUNCA se corten los números ──
  const dlFont = {
    barPrev: (ctx) => {
      const w = ctx.chart.width || 800;
      const count = ctx.chart.data?.labels?.length || 12;
      const sz = Math.max(9, Math.min(12, (w / count) * 0.155));
      return { weight: '700', size: sz, family: "'Inter', sans-serif" };
    },
    barCurr: (ctx) => {
      const w = ctx.chart.width || 800;
      const count = ctx.chart.data?.labels?.length || 12;
      const sz = Math.max(9.5, Math.min(12.5, (w / count) * 0.165));
      return { weight: '800', size: sz, family: "'Inter', sans-serif" };
    },
    plan: (ctx) => {
      const w = ctx.chart.width || 800;
      const count = ctx.chart.data?.labels?.length || 12;
      const sz = Math.max(9, Math.min(11.5, (w / count) * 0.145));
      return { weight: '700', size: sz, family: "'Inter', sans-serif" };
    },
    areaPrev: (ctx) => ({
      weight: '700',
      size: 11,
      family: "'Inter', sans-serif"
    }),
    areaCurr: (ctx) => ({
      weight: '800',
      size: 11.5,
      family: "'Inter', sans-serif"
    })
  };

  function shouldShowDataLabel(ctx, userShowLabels) {
    if (!userShowLabels) return false;
    const chartW = ctx.chart.width || 400;
    const labelCount = ctx.chart.data?.labels?.length || 1;
    const pxPerWeek = chartW / labelCount;
    // Si hay menos de 28px de ancho por semana en la pantalla (como 40 semanas en móvil o tablet),
    // ocultamos los rótulos automáticos para que no se empasten ni tapen las barras.
    // El usuario siempre puede tocar cualquier barra para ver el detalle en el tooltip interactivo.
    return pxPerWeek >= 28;
  }

  function getBaseChartOptions(showLabels, currentYear = 2026) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: { top: 4, bottom: 6, left: 8, right: 8 }
      },
      plugins: {
        legend: {
          position: 'top',
          align: 'center',
          labels: {
            usePointStyle: true,
            boxWidth: 9,
            padding: 10,
            font: { size: 11.5, weight: '600', family: "'Inter', sans-serif" },
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
            title: (items) => {
              if (!items || !items.length) return '';
              const rawLabel = items[0].label;
              const weekLabel = Array.isArray(rawLabel) ? rawLabel[0] : rawLabel;
              const wMatch = String(weekLabel).match(/S(\d+)/i);
              if (wMatch) {
                const w = parseInt(wMatch[1], 10);
                const m = getMonthForWeek(w, currentYear);
                const mName = MONTH_NAMES_FULL[m] || '';
                return `Semana ${w} · ${mName} ${currentYear}`;
              }
              return String(weekLabel);
            },
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
          display: (c) => shouldShowDataLabel(c, showLabels),
          clip: false,
          clamp: false
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            maxRotation: 0,
            autoSkip: true,
            font: { weight: '700', size: 10.5, family: "'Inter', sans-serif" },
            color: '#475569'
          }
        },
        y: {
          beginAtZero: true,
          grace: '10%',
          grid: { color: 'rgba(226, 232, 240, 0.6)' },
          ticks: {
            font: { size: 10.5, family: "'Inter', sans-serif" },
            color: '#64748b',
            callback: (v) => formatNumberBadge(v)
          }
        }
      },
      interaction: { mode: 'index', intersect: false }
    };
  }

  function renderBarChart(canvasId, labels, barSeries, planLine, showLabels, currentYear = 2026) {
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
        barPercentage: 0.90,
        categoryPercentage: 0.88,
        datalabels: {
          display: (c) => shouldShowDataLabel(c, showLabels),
          clip: false,
          clamp: false,
          anchor: (c) => {
            const val = c.dataset.data[c.dataIndex];
            const yAxis = c.chart.scales.y;
            if (!val || val === 0 || !yAxis) return 'center';
            const h = Math.abs(yAxis.getPixelForValue(0) - yAxis.getPixelForValue(val));
            return h < 26 ? 'end' : 'center';
          },
          align: (c) => {
            const val = c.dataset.data[c.dataIndex];
            const yAxis = c.chart.scales.y;
            if (!val || val === 0 || !yAxis) return 'center';
            const h = Math.abs(yAxis.getPixelForValue(0) - yAxis.getPixelForValue(val));
            return h < 26 ? 'top' : 'center';
          },
          offset: (c) => {
            const val = c.dataset.data[c.dataIndex];
            const yAxis = c.chart.scales.y;
            if (!val || val === 0 || !yAxis) return 0;
            const h = Math.abs(yAxis.getPixelForValue(0) - yAxis.getPixelForValue(val));
            return h < 26 ? 3 : 0;
          },
          color: '#0f172a',
          font: dlFont.barPrev,
          textStrokeColor: 'rgba(255, 255, 255, 0.95)',
          textStrokeWidth: 2,
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
        barPercentage: 0.90,
        categoryPercentage: 0.88,
        datalabels: {
          display: (c) => shouldShowDataLabel(c, showLabels),
          clip: false,
          clamp: false,
          anchor: (c) => {
            const val = c.dataset.data[c.dataIndex];
            const yAxis = c.chart.scales.y;
            if (!val || val === 0 || !yAxis) return 'center';
            const h = Math.abs(yAxis.getPixelForValue(0) - yAxis.getPixelForValue(val));
            return h < 26 ? 'end' : 'center';
          },
          align: (c) => {
            const val = c.dataset.data[c.dataIndex];
            const yAxis = c.chart.scales.y;
            if (!val || val === 0 || !yAxis) return 'center';
            const h = Math.abs(yAxis.getPixelForValue(0) - yAxis.getPixelForValue(val));
            return h < 26 ? 'top' : 'center';
          },
          offset: (c) => {
            const val = c.dataset.data[c.dataIndex];
            const yAxis = c.chart.scales.y;
            if (!val || val === 0 || !yAxis) return 0;
            const h = Math.abs(yAxis.getPixelForValue(0) - yAxis.getPixelForValue(val));
            return h < 26 ? 3 : 0;
          },
          color: '#ffffff',
          font: dlFont.barCurr,
          textStrokeColor: 'rgba(15, 23, 42, 0.75)',
          textStrokeWidth: 1.5,
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
          display: (c) => shouldShowDataLabel(c, showLabels),
          clip: false,
          clamp: false,
          align: 'top',
          anchor: 'end',
          offset: 5,
          backgroundColor: 'rgba(255, 255, 255, 0.96)',
          borderColor: planLine.color,
          borderWidth: 1.2,
          color: planLine.color,
          font: dlFont.plan,
          borderRadius: 4,
          padding: { top: 2, bottom: 2, left: 4, right: 4 },
          formatter: (v) => (v !== null && v > 0) ? `P: ${formatNumberBadge(v)}` : ''
        }
      });
    }

    charts[canvasId] = new Chart(ctx, {
      data: { labels, datasets },
      options: getBaseChartOptions(showLabels, currentYear)
    });
  }

  function renderAreaChart(canvasId, labels, areaSeries, planLine, showLabels, currentYear = 2026) {
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
          display: (c) => shouldShowDataLabel(c, showLabels),
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
          display: (c) => shouldShowDataLabel(c, showLabels),
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
          display: (c) => shouldShowDataLabel(c, showLabels),
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
      options: getBaseChartOptions(showLabels, currentYear)
    });
  }

  // ══════════════════════════════════════════════
  // TABLE RENDERER WITH YOY BADGES & PLAN PROJECTION
  // ══════════════════════════════════════════════
  function renderTable(prefix, weekNumbers, currentYear, prevYear, dataMap, lastDataWeek) {
    const container = document.getElementById(`table${prefix}Container`);
    if (!container) return;

    const isFiltering = selectedDivisions.size > 0 && selectedDivisions.size < 12;
    const sortedCodes = Array.from(selectedDivisions).sort();
    const divTableTag = isFiltering
      ? ` · Filtrando ${sortedCodes.length} ${sortedCodes.length === 1 ? 'División' : 'Divisiones'} (${sortedCodes.join(', ')})`
      : '';

    let html = `
      <div class="table-card-header">
        <div class="table-wrapper-title">
          <i class="fa-solid fa-table-list text-primary"></i>
          <span>Detalle Corporativo Semanal (${currentYear} vs ${prevYear}) · Últimas ${weekNumbers.length} Semanas Cerradas${divTableTag}</span>
        </div>
        <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
          <button class="btn-table-copy" onclick="window.copyCorporateTableImage('${prefix}')" title="Copiar tabla desde 'Área' como imagen para PowerPoint">
            <i class="fa-solid fa-camera"></i> Copiar Imagen (desde Área)
          </button>
          <span class="table-scroll-hint">
            <i class="fa-solid fa-arrows-left-right text-primary"></i> Desliza para ver más semanas
          </span>
        </div>
      </div>
      <div class="table-scroll-wrapper">
        <table class="data-table" id="tableCorporateWeekly_${prefix}">
          <thead>
            <tr>
              <th class="col-sticky">Área</th>
    `;
    
    weekNumbers.forEach(w => {
      const mIdx = getMonthForWeek(w, currentYear);
      const mShort = MONTH_NAMES_SHORT[mIdx] || '';
      const mFull = MONTH_NAMES_FULL[mIdx] || '';
      html += `
        <th title="Semana ${w} · ${mFull} ${currentYear}">
          <div class="th-week">S${w}</div>
          <div class="th-month">${mShort}</div>
        </th>`;
    });
    html += `<th class="col-promedio">Promedio</th></tr></thead><tbody>`;

    const metrics = [
      { key: 'recibo', planKey: 'planRecibo', label: '📦 RECIBO', color: '#2563eb' },
      { key: 'despacho', planKey: 'planDespacho', label: '🚛 DESPACHO', color: '#ea580c' },
      { key: 'inventario', planKey: 'planInv', label: '📊 INVENTARIO', color: '#059669' }
    ];

    metrics.forEach(m => {
      // 1. Fila Actual (Real) - Promedio Semanal
      html += `<tr><td class="col-sticky" style="color:${m.color}; font-weight:700;" title="${m.label} (Real)">${m.label}</td>`;
      let sumCurr = 0;
      let countCurr = 0;
      weekNumbers.forEach(w => {
        const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][m.key] : 0;
        sumCurr += val;
        if (val > 0) countCurr++;
        html += `<td style="font-weight:600;">${Math.round(val).toLocaleString('es-PE')}</td>`;
      });
      const avgCurr = countCurr > 0 ? sumCurr / countCurr : (weekNumbers.length > 0 ? sumCurr / weekNumbers.length : 0);
      html += `<td class="col-promedio" style="font-weight:800; background:#f1f5f9;">${Math.round(avgCurr).toLocaleString('es-PE')}</td></tr>`;

      // 2. Fila Plan (Metas del período cerrado) - Promedio Semanal
      html += `<tr style="color:#0284c7; background:rgba(239, 246, 255, 0.35);"><td class="col-sticky" style="font-weight:600; padding-left: 14px; color:#0284c7; background:#eff6ff;" title="Plan Objetivo ${currentYear}">└ Plan Objetivo</td>`;
      let sumPlan = 0;
      let countPlan = 0;
      weekNumbers.forEach(w => {
        const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][m.planKey] : 0;
        sumPlan += val;
        if (val > 0) countPlan++;
        html += `<td>${val > 0 ? Math.round(val).toLocaleString('es-PE') : '--'}</td>`;
      });
      const avgPlan = countPlan > 0 ? sumPlan / countPlan : 0;
      html += `<td class="col-promedio" style="font-weight:700; background:#eff6ff;">${avgPlan > 0 ? Math.round(avgPlan).toLocaleString('es-PE') : '--'}</td></tr>`;

      // 3. Fila Año Anterior (Real) - Promedio Semanal
      html += `<tr style="color:#64748b;"><td class="col-sticky" style="font-weight:600; padding-left: 14px; color:#64748b; background:#f8fafc;" title="Real ${prevYear}">└ Real ${prevYear}</td>`;
      let sumPrev = 0;
      let countPrev = 0;
      weekNumbers.forEach(w => {
        const val = dataMap[`${prevYear}-${w}`] ? dataMap[`${prevYear}-${w}`][m.key] : 0;
        sumPrev += val;
        if (val > 0) countPrev++;
        html += `<td>${val > 0 ? Math.round(val).toLocaleString('es-PE') : '--'}</td>`;
      });
      const avgPrev = countPrev > 0 ? sumPrev / countPrev : 0;
      html += `<td class="col-promedio" style="font-weight:700; background:#f8fafc;">${avgPrev > 0 ? Math.round(avgPrev).toLocaleString('es-PE') : '--'}</td></tr>`;
    });

    html += '</tbody></table></div>';
    container.innerHTML = html;
  }

  // ══════════════════════════════════════════════
  // TABLA 2: MOVIMIENTOS POR DIVISIÓN (AÑO ACTUAL) CON FLECHAS DE TENDENCIA
  // ══════════════════════════════════════════════
  let divTableProcess = { Secos: 'all', Frescos: 'all' };

  window.setDivTableProcessView = function(prefix, proc) {
    divTableProcess[prefix] = proc;
    renderAll();
  };

  function getDivVal(prefix, code, proc, weekKey, dataMap) {
    // 1. Verificar si existen datos parseados directos
    const divExact = window.parsedDivisionData?.[prefix]?.[code]?.[proc];
    if (divExact && divExact[weekKey] !== undefined) {
      return parseFloat(divExact[weekKey]) || 0;
    }
    // 2. Verificar datos predeterminados verificados
    const defVal = DEFAULT_DIVISIONS_DATA?.[prefix]?.[code]?.[proc]?.[weekKey];
    if (defVal !== undefined) {
      return parseFloat(defVal) || 0;
    }
    // 3. Fallback: Participación porcentual de la división sobre el total corporativo
    if (dataMap && dataMap[weekKey]) {
      const totalVal = dataMap[weekKey][proc] || 0;
      const share = DIVISION_SHARE[prefix]?.[code] || 0;
      return totalVal * share;
    }
    return 0;
  }

  function renderDivisionsTable(prefix, weekNumbers, currentYear, dataMap, lastDataWeek) {
    const container = document.getElementById(`tableDivisions${prefix}Container`);
    if (!container) return;

    const currentView = divTableProcess[prefix] || 'all';
    const isFiltering = selectedDivisions.size > 0 && selectedDivisions.size < 12;

    // Divisiones correspondientes a cada Centro de Distribución
    const defaultDivs = prefix === 'Frescos'
      ? ['J01', 'J03', 'J04', 'J05', 'J06', 'J07']
      : ['J01', 'J02', 'J05', 'J06', 'J07', 'J08', 'J09', 'J10', 'J11', 'J12'];

    let divCodes = [...defaultDivs];
    if (isFiltering) {
      const activeInCd = defaultDivs.filter(c => selectedDivisions.has(c));
      if (activeInCd.length > 0) {
        divCodes = activeInCd;
      }
    }

    // Semanas de comparación de tendencia: última cerrada vs anterior a la cerrada
    const W_last = weekNumbers.length > 0 ? weekNumbers[weekNumbers.length - 1] : lastDataWeek;
    const W_prev = weekNumbers.length > 1 ? weekNumbers[weekNumbers.length - 2] : Math.max(1, W_last - 1);

    const procConfigs = [
      { key: 'recibo', title: 'Recibo (Entradas)', icon: '📦', color: '#2563eb', cardCls: 'division-process-recibo' },
      { key: 'despacho', title: 'Despacho (Salidas)', icon: '🚛', color: '#ea580c', cardCls: 'division-process-despacho' },
      { key: 'inventario', title: 'Inventario (Stock)', icon: '📊', color: '#059669', cardCls: 'division-process-inventario' }
    ];

    const procsToRender = currentView === 'all'
      ? procConfigs
      : procConfigs.filter(p => p.key === currentView);

    let topHeaderHtml = `
      <div class="table-card-header" style="flex-wrap:wrap; gap:12px; margin-bottom:14px;">
        <div class="table-wrapper-title">
          <i class="fa-solid fa-boxes-stacked text-primary"></i>
          <span>Movimiento de Cajas por División · Año ${currentYear} · Últimas ${weekNumbers.length} Semanas Cerradas</span>
          ${isFiltering ? `<span style="font-size:0.75rem; background:#eff6ff; color:#1d4ed8; padding:2px 8px; border-radius:12px; border:1px solid #bfdbfe; font-weight:700;">Filtrando ${divCodes.length} div.</span>` : ''}
        </div>
        <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
          <button class="btn-table-copy" onclick="window.copyActiveDivisionTableImage('${prefix}')" title="Copiar tabla de división como imagen para PowerPoint">
            <i class="fa-solid fa-camera"></i> Copiar Imagen
          </button>
          <div class="division-view-pills">
            <button class="division-view-pill ${currentView === 'all' ? 'active' : ''}" onclick="window.setDivTableProcessView('${prefix}', 'all')">
              <i class="fa-solid fa-table-columns"></i> Todos (3 en 1)
            </button>
            <button class="division-view-pill ${currentView === 'recibo' ? 'active' : ''}" onclick="window.setDivTableProcessView('${prefix}', 'recibo')">
              📦 Recibo
            </button>
            <button class="division-view-pill ${currentView === 'despacho' ? 'active' : ''}" onclick="window.setDivTableProcessView('${prefix}', 'despacho')">
              🚛 Despacho
            </button>
            <button class="division-view-pill ${currentView === 'inventario' ? 'active' : ''}" onclick="window.setDivTableProcessView('${prefix}', 'inventario')">
              📊 Inventario
            </button>
          </div>
          <span class="table-scroll-hint">
            <i class="fa-solid fa-arrows-left-right text-primary"></i> Desliza para ver más semanas
          </span>
        </div>
      </div>
    `;

    let tablesHtml = '';
    if (currentView === 'all') {
      tablesHtml += `<div class="divisions-multi-table-wrap">`;
    }

    procsToRender.forEach(proc => {
      const cardId = `divProcCard_${prefix}_${proc.key}`;
      tablesHtml += `
        <div class="division-process-table-card" id="${cardId}" style="${currentView !== 'all' ? 'width:100%;' : ''}">
          <div class="division-process-header ${proc.cardCls}">
            <div style="display:flex; align-items:center; gap:8px;">
              <span>${proc.icon}</span>
              <span>${proc.title} ${currentYear}</span>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
              <span style="font-size:0.72rem; font-weight:700; opacity:0.85;">
                Cajas / Semana
              </span>
              <button class="btn-chart-copy" onclick="window.copyDivisionProcessCardImage('${cardId}', '${proc.title} ${currentYear}')" title="Copiar esta tabla de ${proc.title} como imagen para PowerPoint">
                <i class="fa-solid fa-camera"></i> Copiar Imagen
              </button>
            </div>
          </div>
          <div class="table-scroll-wrapper">
            <table class="data-table">
              <thead>
                <tr>
                  <th class="col-sticky col-division-code">Div.</th>
      `;

      weekNumbers.forEach(w => {
        const mIdx = getMonthForWeek(w, currentYear);
        const mShort = MONTH_NAMES_SHORT[mIdx] || '';
        const isLast = (w === W_last);
        tablesHtml += `
          <th title="Semana ${w}" ${isLast ? 'style="background:#eff6ff;"' : ''}>
            <div class="th-week" ${isLast ? 'style="color:#2563eb;"' : ''}>S${w}</div>
            <div class="th-month">${mShort}</div>
          </th>`;
      });

      tablesHtml += `
                  <th style="min-width:115px; background:#f8fafc;" title="Tendencia comparativa semana ${W_last} vs semana ${W_prev}">
                    <div>Tendencia</div>
                    <div style="font-size:0.68rem; font-weight:700; color:#64748b;">S${W_last} vs S${W_prev}</div>
                  </th>
                </tr>
              </thead>
              <tbody>
      `;

      // Acumuladores de totales por semana
      const weekTotals = {};
      weekNumbers.forEach(w => { weekTotals[w] = 0; });
      let totalAllWeeks = 0;

      divCodes.forEach(code => {
        const fullName = DIVISION_NAMES[code] || '';
        tablesHtml += `<tr>`;
        tablesHtml += `
          <td class="col-sticky col-division-code" title="${code} - ${fullName}">
            ${code}
          </td>
        `;

        let sumDiv = 0;
        let countDiv = 0;

        weekNumbers.forEach(w => {
          const val = getDivVal(prefix, code, proc.key, `${currentYear}-${w}`, dataMap);
          weekTotals[w] += val;
          sumDiv += val;
          if (val > 0) countDiv++;
          const isLastCol = (w === W_last);
          tablesHtml += `<td style="font-weight:600; ${isLastCol ? 'background:rgba(239, 246, 255, 0.4);' : ''}">${val > 0 ? Math.round(val).toLocaleString('es-PE') : '-'}</td>`;
        });

        // Tendencia: Comparación W_last vs W_prev
        const vLast = getDivVal(prefix, code, proc.key, `${currentYear}-${W_last}`, dataMap);
        const vPrev = getDivVal(prefix, code, proc.key, `${currentYear}-${W_prev}`, dataMap);
        const diff = vLast - vPrev;
        let trendHtml = '';

        if (vPrev > 0) {
          const pct = ((vLast - vPrev) / vPrev) * 100;
          if (diff >= 0) {
            trendHtml = `
              <span class="trend-arrow trend-up" title="S${W_last} (${Math.round(vLast).toLocaleString('es-PE')}) vs S${W_prev} (${Math.round(vPrev).toLocaleString('es-PE')}): Aumentó +${Math.round(diff).toLocaleString('es-PE')} cajas (+${pct.toFixed(1)}%)">
                <i class="fa-solid fa-arrow-up"></i> ▲ +${pct.toFixed(1)}%
              </span>`;
          } else {
            trendHtml = `
              <span class="trend-arrow trend-down" title="S${W_last} (${Math.round(vLast).toLocaleString('es-PE')}) vs S${W_prev} (${Math.round(vPrev).toLocaleString('es-PE')}): Disminuyó ${Math.round(diff).toLocaleString('es-PE')} cajas (${pct.toFixed(1)}%)">
                <i class="fa-solid fa-arrow-down"></i> ▼ ${pct.toFixed(1)}%
              </span>`;
          }
        } else if (vPrev === 0 && vLast > 0) {
          trendHtml = `<span class="trend-arrow trend-up" title="S${W_last}: Nuevo volumen (+${Math.round(vLast).toLocaleString('es-PE')})">▲ Nuevo</span>`;
        } else {
          trendHtml = `<span style="color:#94a3b8; font-size:0.75rem;">--</span>`;
        }

        const avgDiv = countDiv > 0 ? sumDiv / countDiv : 0;
        totalAllWeeks += sumDiv;

        tablesHtml += `
          <td style="text-align:center;">${trendHtml}</td>
        </tr>`;
      });

      // Fila de Total por Proceso
      const totLast = weekTotals[W_last] || 0;
      const totPrev = weekTotals[W_prev] || 0;
      const totDiff = totLast - totPrev;
      let totTrendHtml = '';

      if (totPrev > 0) {
        const totPct = ((totLast - totPrev) / totPrev) * 100;
        if (totDiff >= 0) {
          totTrendHtml = `
            <span class="trend-arrow trend-up" title="Total S${W_last} (${Math.round(totLast).toLocaleString('es-PE')}) vs Total S${W_prev} (${Math.round(totPrev).toLocaleString('es-PE')}): +${Math.round(totDiff).toLocaleString('es-PE')} cajas (+${totPct.toFixed(1)}%)">
              <i class="fa-solid fa-arrow-up"></i> ▲ +${totPct.toFixed(1)}%
            </span>`;
        } else {
          totTrendHtml = `
            <span class="trend-arrow trend-down" title="Total S${W_last} (${Math.round(totLast).toLocaleString('es-PE')}) vs Total S${W_prev} (${Math.round(totPrev).toLocaleString('es-PE')}): ${Math.round(totDiff).toLocaleString('es-PE')} cajas (${totPct.toFixed(1)}%)">
              <i class="fa-solid fa-arrow-down"></i> ▼ ${totPct.toFixed(1)}%
            </span>`;
        }
      } else {
        totTrendHtml = `<span style="color:#94a3b8; font-size:0.75rem;">--</span>`;
      }

      tablesHtml += `
        <tr class="row-total">
          <td class="col-sticky col-division-code" style="font-weight:800; color:#0f172a;">
            Total
          </td>
      `;
      weekNumbers.forEach(w => {
        const isLastCol = (w === W_last);
        tablesHtml += `<td style="font-weight:800; ${isLastCol ? 'background:#e0f2fe; color:#0369a1;' : ''}">${Math.round(weekTotals[w]).toLocaleString('es-PE')}</td>`;
      });

      tablesHtml += `
          <td style="text-align:center;">${totTrendHtml}</td>
        </tr>
      `;

      tablesHtml += `
              </tbody>
            </table>
          </div>
        </div>
      `;
    });

    if (currentView === 'all') {
      tablesHtml += `</div>`;
    }

    // Glosario de Divisiones Corporativas a un Costado
    let glossaryHtml = `
      <div class="divisions-glossary-card" id="glossaryCard_${prefix}">
        <div class="glossary-header">
          <div style="display:flex; align-items:center; gap:8px;">
            <i class="fa-solid fa-book-bookmark text-primary"></i>
            <span>Glosario Divisiones</span>
          </div>
          <button class="btn-table-copy" onclick="window.copyGlossaryImage('${prefix}')" title="Copiar Glosario como imagen para PowerPoint" style="padding:4px 9px; font-size:0.75rem;">
            <i class="fa-solid fa-camera"></i> Copiar
          </button>
        </div>
        <div class="glossary-body">
          <table class="glossary-table">
            <thead>
              <tr>
                <th style="width:50px; text-align:center;">Cód.</th>
                <th>Nombre Oficial</th>
              </tr>
            </thead>
            <tbody>
    `;

    defaultDivs.forEach(c => {
      const name = DIVISION_NAMES[c] || '';
      const isSel = selectedDivisions.has(c);
      glossaryHtml += `
        <tr class="${isSel ? 'glossary-row-highlight' : ''}">
          <td style="text-align:center;">
            <span class="div-code-badge" style="font-size:0.75rem; padding:2px 6px;">${c}</span>
          </td>
          <td style="font-weight:600; font-size:0.8rem; color:#1e293b;">
            ${name}
          </td>
        </tr>
      `;
    });

    glossaryHtml += `
            </tbody>
          </table>
        </div>
      </div>
    `;

    container.innerHTML = `
      ${topHeaderHtml}
      <div class="division-section-content">
        <div class="division-tables-main">
          ${tablesHtml}
        </div>
        <div class="division-glossary-sidebar">
          ${glossaryHtml}
        </div>
      </div>
    `;
  }

  // ══════════════════════════════════════════════
  // NOTIFICACIONES FLOTANTES (TOAST)
  // ══════════════════════════════════════════════
  function showToast(msg, type = 'info', duration = 3500) {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    const icon = type === 'success' ? 'fa-circle-check' : (type === 'danger' ? 'fa-triangle-exclamation' : 'fa-circle-info');
    toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${msg}</span>`;
    container.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => toast.remove(), 300);
    }, duration);
  }

  // ══════════════════════════════════════════════
  // LIMPIEZA Y PARSEO ROBUSTO DE NÚMEROS (ES/EN)
  // ══════════════════════════════════════════════
  function cleanNumber(val) {
    if (val === null || val === undefined) return 0;
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    let str = String(val).trim().replace(/[^\d.,\-]/g, '');
    if (!str) return 0;
    if (str.includes('.') && str.includes(',')) {
      if (str.indexOf('.') < str.indexOf(',')) {
        str = str.replace(/\./g, '').replace(',', '.');
      } else {
        str = str.replace(/,/g, '');
      }
    } else if (str.includes(',')) {
      const parts = str.split(',');
      if (parts.length === 2 && parts[1].length <= 2) {
        str = str.replace(',', '.');
      } else {
        str = str.replace(/,/g, '');
      }
    } else if (str.includes('.')) {
      const parts = str.split('.');
      if (parts.length === 2 && parts[1].length === 3) {
        str = str.replace(/\./g, '');
      } else if (parts.length > 2) {
        str = str.replace(/\./g, '');
      }
    }
    const n = parseFloat(str);
    return isNaN(n) ? 0 : n;
  }

  // ══════════════════════════════════════════════
  // BOTONES PARA COPIAR INFORMACIÓN DE THROUGHPUT (EXCEL / REPORTES)
  // ══════════════════════════════════════════════
  function fallbackClipboardCopy(text, onSuccess, onError) {
    try {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.style.position = 'fixed';
      textarea.style.top = '0';
      textarea.style.left = '0';
      textarea.style.width = '2em';
      textarea.style.height = '2em';
      textarea.style.padding = '0';
      textarea.style.border = 'none';
      textarea.style.outline = 'none';
      textarea.style.boxShadow = 'none';
      textarea.style.background = 'transparent';
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      const successful = document.execCommand('copy');
      document.body.removeChild(textarea);
      if (successful) {
        if (onSuccess) onSuccess();
      } else {
        if (onError) onError();
      }
    } catch (err) {
      if (onError) onError(err);
    }
  }

  window.copyThroughputProcess = function(processKey) {
    const isFrescos = document.getElementById('tab-frescos')?.classList.contains('active');
    const prefix = isFrescos ? 'Frescos' : 'Secos';
    const cdData = isFrescos ? window.dataFrescos : window.dataSecos;

    if (!cdData || !cdData.rows || cdData.rows.length === 0) {
      showToast('No hay datos disponibles para copiar', 'danger');
      return;
    }

    const cols = resolveColumns(cdData.headers, cdData.rows);
    const procNames = {
      recibo: { name: 'RECIBO (ENTRADAS)', col: cols.recibo, planCol: cols.planRecibo, icon: '📦' },
      despacho: { name: 'DESPACHO (SALIDAS)', col: cols.despacho, planCol: cols.planDespacho, icon: '🚛' },
      inventario: { name: 'INVENTARIO (STOCK)', col: cols.inventario, planCol: cols.planInv, icon: '📊' }
    };
    const procInfo = procNames[processKey] || procNames.recibo;

    // Obtener semanas cerradas
    const allParsed = [];
    cdData.rows.forEach(row => {
      const label = row[cols.timeCol];
      if (!label) return;
      const parsed = parseWeekLabel(label);
      if (!parsed || isNaN(parsed.week) || isNaN(parsed.year)) return;
      allParsed.push({
        label,
        week: parsed.week,
        year: parsed.year,
        recibo: parseFloat(row[cols.recibo]) || 0,
        despacho: parseFloat(row[cols.despacho]) || 0,
        inventario: parseFloat(row[cols.inventario]) || 0,
        planRecibo: parseFloat(row[cols.planRecibo]) || 0,
        planDespacho: parseFloat(row[cols.planDespacho]) || 0,
        planInv: parseFloat(row[cols.planInv]) || 0
      });
    });

    const currentYear = Math.max(...allParsed.map(r => r.year));
    const prevYear = currentYear - 1;
    const currentYearRows = allParsed.filter(r => r.year === currentYear).sort((a, b) => a.week - b.week);

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

    const closedWeeks = [];
    const startWeek = Math.max(1, lastDataWeek - 7);
    for (let w = startWeek; w <= lastDataWeek; w++) {
      closedWeeks.push(w);
    }
    const tableWeeks = closedWeeks;
    const W_last = tableWeeks.length > 0 ? tableWeeks[tableWeeks.length - 1] : lastDataWeek;
    const W_prev = tableWeeks.length > 1 ? tableWeeks[tableWeeks.length - 2] : Math.max(1, W_last - 1);

    const dataMap = {};
    allParsed.forEach(r => {
      dataMap[`${r.year}-${r.week}`] = r;
    });

    // Construir texto en formato TSV (Tab-Separated Values) compatible 100% con Excel
    let tsv = `THROUGHPUT CD ${prefix.toUpperCase()} - ${procInfo.name}\tAño ${currentYear} vs ${prevYear}\n`;
    tsv += `Filtro: Últimas ${tableWeeks.length} Semanas Cerradas (S${tableWeeks[0]} a S${tableWeeks[tableWeeks.length - 1]})\n\n`;

    // ── SECCIÓN 1: DETALLE CORPORATIVO SEMANAL ──
    tsv += `1. DETALLE CORPORATIVO SEMANAL\n`;
    tsv += `Concepto\t` + tableWeeks.map(w => `S${w}`).join('\t') + `\tPromedio\n`;

    // Real 2026
    let sumReal = 0, countReal = 0;
    const realVals = tableWeeks.map(w => {
      const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][processKey] : 0;
      sumReal += val;
      if (val > 0) countReal++;
      return Math.round(val);
    });
    const avgReal = countReal > 0 ? Math.round(sumReal / countReal) : 0;
    tsv += `${procInfo.name} ${currentYear}\t` + realVals.join('\t') + `\t${avgReal}\n`;

    // Plan Objetivo 2026
    const planKey = processKey === 'recibo' ? 'planRecibo' : (processKey === 'despacho' ? 'planDespacho' : 'planInv');
    let sumPlan = 0, countPlan = 0;
    const planVals = tableWeeks.map(w => {
      const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][planKey] : 0;
      sumPlan += val;
      if (val > 0) countPlan++;
      return val > 0 ? Math.round(val) : '';
    });
    const avgPlan = countPlan > 0 ? Math.round(sumPlan / countPlan) : '';
    tsv += `Plan Objetivo ${currentYear}\t` + planVals.join('\t') + `\t${avgPlan}\n`;

    // Real 2025
    let sumPrev = 0, countPrev = 0;
    const prevVals = tableWeeks.map(w => {
      const val = dataMap[`${prevYear}-${w}`] ? dataMap[`${prevYear}-${w}`][processKey] : 0;
      sumPrev += val;
      if (val > 0) countPrev++;
      return val > 0 ? Math.round(val) : '';
    });
    const avgPrev = countPrev > 0 ? Math.round(sumPrev / countPrev) : '';
    tsv += `Real ${prevYear}\t` + prevVals.join('\t') + `\t${avgPrev}\n\n`;

    // ── SECCIÓN 2: MOVIMIENTOS POR DIVISIÓN (AÑO ACTUAL) ──
    tsv += `2. MOVIMIENTOS POR DIVISIÓN - ${procInfo.name} ${currentYear}\n`;
    tsv += `División\tNombre Completo\t` + tableWeeks.map(w => `S${w}`).join('\t') + `\tTendencia (S${W_last} vs S${W_prev})\tPromedio\n`;

    const defaultDivs = prefix === 'Frescos'
      ? ['J01', 'J03', 'J04', 'J05', 'J06', 'J07']
      : ['J01', 'J02', 'J05', 'J06', 'J07', 'J08', 'J09', 'J10', 'J11', 'J12'];

    const divWeekTotals = {};
    tableWeeks.forEach(w => { divWeekTotals[w] = 0; });
    let divGrandTotal = 0;

    defaultDivs.forEach(code => {
      const name = DIVISION_NAMES[code] || '';
      let sumDiv = 0, countDiv = 0;
      const vals = tableWeeks.map(w => {
        const v = getDivVal(prefix, code, processKey, `${currentYear}-${w}`, dataMap);
        sumDiv += v;
        divWeekTotals[w] += v;
        if (v > 0) countDiv++;
        return v > 0 ? Math.round(v) : 0;
      });

      const vLast = getDivVal(prefix, code, processKey, `${currentYear}-${W_last}`, dataMap);
      const vPrev = getDivVal(prefix, code, processKey, `${currentYear}-${W_prev}`, dataMap);
      let trendStr = '--';
      if (vPrev > 0) {
        const pct = ((vLast - vPrev) / vPrev) * 100;
        trendStr = `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`;
      }

      const avgDiv = countDiv > 0 ? Math.round(sumDiv / countDiv) : 0;
      divGrandTotal += sumDiv;

      tsv += `${code}\t${name}\t` + vals.join('\t') + `\t${trendStr}\t${avgDiv}\n`;
    });

    // Fila Total Divisiones
    const totLast = divWeekTotals[W_last] || 0;
    const totPrev = divWeekTotals[W_prev] || 0;
    let totTrendStr = '--';
    if (totPrev > 0) {
      const totPct = ((totLast - totPrev) / totPrev) * 100;
      totTrendStr = `${totPct >= 0 ? '+' : ''}${totPct.toFixed(1)}%`;
    }
    const avgGrand = tableWeeks.length > 0 ? Math.round(divGrandTotal / tableWeeks.length) : 0;
    tsv += `TOTAL\tTotal ${procInfo.name}\t` + tableWeeks.map(w => Math.round(divWeekTotals[w] || 0)).join('\t') + `\t${totTrendStr}\t${avgGrand}\n`;

    // Función de éxito para copiar
    const onCopySuccess = () => {
      showToast(`¡Datos de ${procInfo.name} (CD ${prefix.toUpperCase()}) copiados al portapapeles! Listo para pegar en Excel (Ctrl + V).`, 'success', 4000);

      // Feedback visual en botones
      document.querySelectorAll(`.btn-copy-${processKey}, .btn-chart-copy`).forEach(btn => {
        const origHtml = btn.innerHTML;
        btn.innerHTML = `<i class="fa-solid fa-check"></i> ¡Copiado!`;
        btn.style.filter = 'brightness(1.15)';
        setTimeout(() => {
          btn.innerHTML = origHtml;
          btn.style.filter = '';
        }, 1800);
      });
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsv).then(onCopySuccess).catch(() => {
        fallbackClipboardCopy(tsv, onCopySuccess, () => {
          showToast('No se pudo copiar automáticamente. Por favor inténtalo de nuevo.', 'danger');
        });
      });
    } else {
      fallbackClipboardCopy(tsv, onCopySuccess, () => {
        showToast('No se pudo copiar automáticamente. Por favor inténtalo de nuevo.', 'danger');
      });
    }
  };

  // ══════════════════════════════════════════════
  // COPIAR ELEMENTOS (TABLAS Y GRÁFICOS) COMO IMAGEN PARA POWERPOINT
  // ══════════════════════════════════════════════
  let isCopyingImageInProgress = false;

  function roundRect(ctx, x, y, w, h, r, fill = true, stroke = true) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + r);
    ctx.lineTo(x + w, y + h - r);
    ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    ctx.lineTo(x + r, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
    ctx.closePath();
    if (fill) ctx.fill();
    if (stroke) ctx.stroke();
  }

  function ensureHtml2Canvas() {
    return new Promise((resolve, reject) => {
      if (window.html2canvas) return resolve(window.html2canvas);
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js';
      s.onload = () => resolve(window.html2canvas);
      s.onerror = (e) => reject(new Error('No se pudo cargar la librería html2canvas'));
      document.head.appendChild(s);
    });
  }

  // Escribe un blob de imagen PNG en el portapapeles. BAJO NINGUNA CIRCUNSTANCIA DESCARGA ARCHIVOS.
  async function writeBlobToClipboard(blob, labelName) {
    if (!blob) {
      showToast('Error al generar la imagen', 'danger');
      return;
    }
    try {
      if (window.focus) window.focus();
      if (navigator.clipboard && window.ClipboardItem) {
        const item = new ClipboardItem({ 'image/png': blob });
        await navigator.clipboard.write([item]);
        showToast(`¡Imagen de ${labelName} copiada! Lista para pegar en PowerPoint (Ctrl + V)`, 'success', 4500);
      } else {
        showToast('El portapapeles de imágenes no está disponible en este navegador.', 'danger', 4000);
      }
    } catch (clipErr) {
      console.warn('Error al escribir imagen en el portapapeles:', clipErr);
      if (!document.hasFocus()) {
        showToast('⚠️ La ventana del navegador perdió el foco antes de terminar el copiado. Por seguridad de Windows, mantén la ventana abierta durante 1 segundo al presionar copiar.', 'warning', 6000);
      } else {
        showToast('No se pudo copiar la imagen al portapapeles. Vuelve a intentarlo.', 'danger', 4000);
      }
    }
  }

  // 1. Copiar Slide Completo 3 en 1 Panorámico 16:9 (Widescreen 1920x1080) Ultrarrápido (<20ms, sin esperas ni errores de foco)
  window.copyTripleChartsSlideImage = async function() {
    if (isCopyingImageInProgress) {
      showToast('Copiado en proceso, por favor espera un momento...', 'info', 1800);
      return;
    }
    const isFrescos = document.getElementById('tab-frescos')?.classList.contains('active');
    const prefix = isFrescos ? 'Frescos' : 'Secos';

    const cRecibo = document.getElementById(`chart${prefix}Recibo`);
    const cDespacho = document.getElementById(`chart${prefix}Despacho`);
    const cInventario = document.getElementById(`chart${prefix}Inventario`);

    if (!cRecibo || !cDespacho || !cInventario) {
      showToast('No se encontraron los gráficos para la diapositiva', 'danger');
      return;
    }

    isCopyingImageInProgress = true;

    try {
      // Dimensiones Widescreen estándar de PowerPoint 16:9 (1920 x 1080 Full HD)
      const W = 1920;
      const H = 1080;
      const offscreen = document.createElement('canvas');
      offscreen.width = W;
      offscreen.height = H;
      const ctx = offscreen.getContext('2d');

      // 1. Fondo blanco puro
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, W, H);

      // 2. Título principal centrado ejecutiva
      ctx.fillStyle = '#0f172a';
      ctx.font = '800 32px Inter, -apple-system, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(`THROUGHPUT – CD ${prefix.toUpperCase()} 2026`, W / 2, 42);

      // 3. Dibujar los 3 bloques apilados aprovechando todo el ancho de la diapositiva (sin espacios blancos laterales)
      const topMargin = 72;
      const bottomMargin = 20;
      const sideMargin = 35;
      const cardWidth = W - (sideMargin * 2); // 1850px de ancho
      const availableHeight = H - topMargin - bottomMargin; // 988px
      const gap = 14;
      const cardHeight = Math.floor((availableHeight - (gap * 2)) / 3); // ~320px

      const sections = [
        { canvas: cRecibo, title: 'Entradas (Recibo)', color: '#2563eb', badge: 'Cajas / Semana' },
        { canvas: cDespacho, title: 'Salidas (Despacho)', color: '#ea580c', badge: 'Cajas / Semana' },
        { canvas: cInventario, title: 'Inventario Activo', color: '#059669', badge: 'Stock en Cajas · Barras Comparativas' }
      ];

      sections.forEach((sec, idx) => {
        const y0 = topMargin + idx * (cardHeight + gap);

        // Tarjeta contenedor con borde suave
        ctx.fillStyle = '#ffffff';
        ctx.strokeStyle = '#e2e8f0';
        ctx.lineWidth = 1.2;
        roundRect(ctx, sideMargin, y0, cardWidth, cardHeight, 10, true, true);

        // Cabecera de la sección
        const headerY = y0 + 18;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'middle';
        ctx.fillStyle = sec.color;
        ctx.font = '800 15px Inter, -apple-system, sans-serif';
        ctx.fillText(sec.title, sideMargin + 16, headerY);

        // Badge a la derecha
        ctx.textAlign = 'right';
        ctx.fillStyle = '#64748b';
        ctx.font = '700 11.5px Inter, -apple-system, sans-serif';
        ctx.fillText(sec.badge, sideMargin + cardWidth - 16, headerY);

        // Gráfico (Canvas de Chart.js) estirado en todo el ancho para llenar la diapositiva
        const chartX = sideMargin + 10;
        const chartY = y0 + 32;
        const chartW = cardWidth - 20;
        const chartH = cardHeight - 38;

        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(sec.canvas, chartX, chartY, chartW, chartH);
      });

      offscreen.toBlob(async (blob) => {
        await writeBlobToClipboard(blob, `Diapositiva 3 en 1 Panorámica (CD ${prefix.toUpperCase()})`);

        const btnSlide = document.getElementById('btnCopySlideTriple');
        if (btnSlide) {
          const origHtml = btnSlide.innerHTML;
          btnSlide.innerHTML = `<i class="fa-solid fa-check"></i> ¡Diapositiva Copiada!`;
          setTimeout(() => { btnSlide.innerHTML = origHtml; }, 2500);
        }
      }, 'image/png');

    } catch (err) {
      console.error('Error al componer diapositiva 3 en 1:', err);
      showToast('No se pudo generar la diapositiva 3 en 1.', 'danger');
    } finally {
      setTimeout(() => { isCopyingImageInProgress = false; }, 300);
    }
  };

  // 2. Copiar Tabla Detalle Corporativo Semanal (captura estrictamente desde "Área" sin espacio blanco lateral)
  window.copyCorporateTableImage = async function(prefix) {
    if (isCopyingImageInProgress) {
      showToast('Copiado en proceso, por favor espera un momento...', 'info', 1800);
      return;
    }
    const table = document.getElementById(`tableCorporateWeekly_${prefix}`);
    if (!table) {
      showToast('Tabla no encontrada para copiar', 'danger');
      return;
    }

    isCopyingImageInProgress = true;
    showToast('Generando imagen de la tabla (desde Área)...', 'info', 1600);

    const scrollWrapper = table.closest('.table-scroll-wrapper');
    const origScrollLeft = scrollWrapper ? scrollWrapper.scrollLeft : 0;
    if (scrollWrapper) scrollWrapper.scrollLeft = 0;

    // Medir ancho exacto de la tabla para eliminar el espacio blanco lateral
    const targetWidth = Math.ceil(table.getBoundingClientRect().width || table.scrollWidth);

    try {
      await ensureHtml2Canvas();
      const canvas = await window.html2canvas(table, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: false,
        logging: false,
        width: targetWidth,
        windowWidth: targetWidth + 40
      });

      if (scrollWrapper) scrollWrapper.scrollLeft = origScrollLeft;

      canvas.toBlob(async (blob) => {
        await writeBlobToClipboard(blob, `Tabla Detalle Corporativo (${prefix})`);
      }, 'image/png');

    } catch (err) {
      if (scrollWrapper) scrollWrapper.scrollLeft = origScrollLeft;
      console.error('Error al capturar tabla corporativa:', err);
      showToast('No se pudo generar la imagen de la tabla. Inténtalo de nuevo.', 'danger');
    } finally {
      setTimeout(() => { isCopyingImageInProgress = false; }, 300);
    }
  };

  // 3. Copiar Tarjeta de Proceso de División (captura desde "Despacho (Salidas) 2026" ajustada al ancho de la tabla)
  window.copyDivisionProcessCardImage = async function(target, labelTitle) {
    if (isCopyingImageInProgress) {
      showToast('Copiado en proceso, por favor espera un momento...', 'info', 1800);
      return;
    }
    const el = typeof target === 'string' ? document.getElementById(target) : target;
    if (!el) {
      showToast('Tarjeta de proceso no encontrada', 'danger');
      return;
    }

    isCopyingImageInProgress = true;
    showToast(`Generando imagen de ${labelTitle || 'tabla divisiones'}...`, 'info', 1600);

    const scrollWrapper = el.querySelector('.table-scroll-wrapper');
    const origScrollLeft = scrollWrapper ? scrollWrapper.scrollLeft : 0;
    if (scrollWrapper) scrollWrapper.scrollLeft = 0;

    // Medir ancho exacto de la tabla interior para evitar espacio blanco a la derecha
    const innerTable = el.querySelector('table');
    let targetWidth = 0;
    if (innerTable) {
      targetWidth = Math.ceil(innerTable.getBoundingClientRect().width || innerTable.scrollWidth);
    } else {
      targetWidth = Math.ceil(el.scrollWidth || el.offsetWidth);
    }

    // Guardar estilos originales del card
    const origWidth = el.style.width;
    const origMaxWidth = el.style.maxWidth;
    const origFlex = el.style.flex;

    // Ajustar temporalmente el ancho del card al ancho exacto de la tabla
    el.style.width = `${targetWidth}px`;
    el.style.maxWidth = `${targetWidth}px`;
    el.style.flex = 'none';

    // Ocultar botones de copia internos durante la captura
    const ignoreBtns = el.querySelectorAll('.btn-chart-copy, .btn-table-copy');
    ignoreBtns.forEach(b => b.setAttribute('data-html2canvas-ignore', 'true'));

    try {
      await ensureHtml2Canvas();
      const canvas = await window.html2canvas(el, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: false,
        logging: false,
        width: targetWidth,
        windowWidth: targetWidth + 40
      });

      // Restaurar estilos originales
      el.style.width = origWidth;
      el.style.maxWidth = origMaxWidth;
      el.style.flex = origFlex;
      ignoreBtns.forEach(b => b.removeAttribute('data-html2canvas-ignore'));
      if (scrollWrapper) scrollWrapper.scrollLeft = origScrollLeft;

      canvas.toBlob(async (blob) => {
        await writeBlobToClipboard(blob, labelTitle || 'Tabla División');
      }, 'image/png');

    } catch (err) {
      el.style.width = origWidth;
      el.style.maxWidth = origMaxWidth;
      el.style.flex = origFlex;
      ignoreBtns.forEach(b => b.removeAttribute('data-html2canvas-ignore'));
      if (scrollWrapper) scrollWrapper.scrollLeft = origScrollLeft;
      console.error('Error al capturar tabla de división:', err);
      showToast('No se pudo generar la imagen de la tabla. Inténtalo de nuevo.', 'danger');
    } finally {
      setTimeout(() => { isCopyingImageInProgress = false; }, 300);
    }
  };

  // 4. Copiar la tabla de división activa según la vista seleccionada
  window.copyActiveDivisionTableImage = function(prefix) {
    const currentView = divTableProcess[prefix] || 'all';
    if (currentView !== 'all') {
      const cardId = `divProcCard_${prefix}_${currentView}`;
      const titles = { recibo: 'Recibo (Entradas)', despacho: 'Despacho (Salidas)', inventario: 'Inventario' };
      const label = `${titles[currentView] || currentView} 2026`;
      window.copyDivisionProcessCardImage(cardId, label);
    } else {
      // Si la vista está en "Todos (3 en 1)", copiamos la tarjeta de Despacho o el primer proceso visible
      const firstCard = document.querySelector(`#tableDivisions${prefix}Container .division-process-table-card`);
      if (firstCard) {
        window.copyDivisionProcessCardImage(firstCard.id, `Tabla Divisiones - CD ${prefix.toUpperCase()}`);
      } else {
        showToast('No se encontró la tabla de divisiones para copiar', 'danger');
      }
    }
  };

  // 5. Copiar Glosario de Divisiones como imagen para PowerPoint
  window.copyGlossaryImage = async function(prefix) {
    if (isCopyingImageInProgress) {
      showToast('Copiado en proceso, por favor espera un momento...', 'info', 1800);
      return;
    }
    const glossaryCard = document.getElementById(`glossaryCard_${prefix}`);
    if (!glossaryCard) {
      showToast('Glosario no encontrado para copiar', 'danger');
      return;
    }

    isCopyingImageInProgress = true;
    showToast('Copiando Glosario de Divisiones...', 'info', 1500);

    const ignoreBtns = glossaryCard.querySelectorAll('.btn-table-copy');
    ignoreBtns.forEach(b => b.setAttribute('data-html2canvas-ignore', 'true'));

    try {
      await ensureHtml2Canvas();
      const canvas = await window.html2canvas(glossaryCard, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: false,
        logging: false
      });

      ignoreBtns.forEach(b => b.removeAttribute('data-html2canvas-ignore'));

      canvas.toBlob(async (blob) => {
        await writeBlobToClipboard(blob, `Glosario de Divisiones (CD ${prefix.toUpperCase()})`);
      }, 'image/png');

    } catch (err) {
      ignoreBtns.forEach(b => b.removeAttribute('data-html2canvas-ignore'));
      console.error('Error al capturar glosario:', err);
      showToast('No se pudo generar la imagen del glosario. Inténtalo de nuevo.', 'danger');
    } finally {
      setTimeout(() => { isCopyingImageInProgress = false; }, 300);
    }
  };

  // 6. Copiar Gráfico individual Throughput (Recibo, Despacho o Inventario) Ultrarrápido (<15ms)
  window.copyThroughputChartImage = async function(processKey) {
    if (isCopyingImageInProgress) {
      showToast('Copiado en proceso, por favor espera un momento...', 'info', 1800);
      return;
    }
    const isFrescos = document.getElementById('tab-frescos')?.classList.contains('active');
    const prefix = isFrescos ? 'Frescos' : 'Secos';

    const canvasIdMap = {
      recibo: `chart${prefix}Recibo`,
      despacho: `chart${prefix}Despacho`,
      inventario: `chart${prefix}Inventario`
    };
    const titleMap = {
      recibo: `Recibo (Entradas) - CD ${prefix.toUpperCase()}`,
      despacho: `Despacho (Salidas) - CD ${prefix.toUpperCase()}`,
      inventario: `Inventario - CD ${prefix.toUpperCase()}`
    };
    const colorMap = {
      recibo: '#2563eb',
      despacho: '#ea580c',
      inventario: '#059669'
    };

    const targetCanvasId = canvasIdMap[processKey] || canvasIdMap.recibo;
    const label = titleMap[processKey] || 'Gráfico Throughput';
    const srcCanvas = document.getElementById(targetCanvasId);
    if (!srcCanvas) {
      showToast('Gráfico no encontrado para copiar', 'danger');
      return;
    }

    isCopyingImageInProgress = true;

    try {
      // Crear canvas panorámico individual (1600 x 720)
      const W = 1600;
      const H = 720;
      const offscreen = document.createElement('canvas');
      offscreen.width = W;
      offscreen.height = H;
      const ctx = offscreen.getContext('2d');

      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, W, H);

      // Card con borde suave
      ctx.fillStyle = '#ffffff';
      ctx.strokeStyle = '#e2e8f0';
      ctx.lineWidth = 1.2;
      roundRect(ctx, 20, 20, W - 40, H - 40, 12, true, true);

      // Cabecera
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      ctx.fillStyle = colorMap[processKey] || '#1e293b';
      ctx.font = '800 20px Inter, -apple-system, sans-serif';
      ctx.fillText(label, 42, 50);

      // Badge
      ctx.textAlign = 'right';
      ctx.fillStyle = '#64748b';
      ctx.font = '700 13px Inter, -apple-system, sans-serif';
      ctx.fillText('Cajas / Semana', W - 42, 50);

      // Dibujar gráfico con suavizado de alta calidad
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(srcCanvas, 30, 75, W - 60, H - 95);

      offscreen.toBlob(async (blob) => {
        await writeBlobToClipboard(blob, label);
      }, 'image/png');

    } catch (err) {
      console.error('Error al capturar gráfico individual:', err);
      showToast('No se pudo copiar el gráfico. Inténtalo de nuevo.', 'danger');
    } finally {
      setTimeout(() => { isCopyingImageInProgress = false; }, 300);
    }
  };

  // 7. Función genérica de respaldo (sin descargas automáticas)
  window.copyElementAsImage = async function(target, labelName) {
    if (isCopyingImageInProgress) return;
    const el = typeof target === 'string' ? document.getElementById(target) : target;
    if (!el) return;

    isCopyingImageInProgress = true;
    showToast(`Generando imagen de ${labelName || 'elemento'}...`, 'info', 1600);

    const actionElements = el.querySelectorAll('.btn-table-copy, .btn-chart-copy, .table-scroll-hint');
    actionElements.forEach(btn => btn.setAttribute('data-html2canvas-ignore', 'true'));

    try {
      await ensureHtml2Canvas();
      const canvas = await window.html2canvas(el, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: false,
        logging: false
      });

      actionElements.forEach(btn => btn.removeAttribute('data-html2canvas-ignore'));

      canvas.toBlob(async (blob) => {
        await writeBlobToClipboard(blob, labelName || 'Elemento');
      }, 'image/png');

    } catch (err) {
      actionElements.forEach(btn => btn.removeAttribute('data-html2canvas-ignore'));
      console.error('Error al capturar elemento:', err);
      showToast('No se pudo generar la imagen. Inténtalo de nuevo.', 'danger');
    } finally {
      setTimeout(() => { isCopyingImageInProgress = false; }, 300);
    }
  };

  // ══════════════════════════════════════════════
  // DATOS PREDETERMINADOS (DEL ARCHIVO EXCEL DE LA EMPRESA)
  // ══════════════════════════════════════════════
  const DEFAULT_DIVISIONS_DATA = {
    Secos: {
      J01: {
        recibo: { '2026-28': 420000, '2026-29': 413783, '2026-30': 398776, '2026-31': 367851, '2026-32': 392610, '2026-33': 428478, '2026-34': 513906, '2026-35': 527438 },
        despacho: { '2026-28': 430000, '2026-29': 419499.83, '2026-30': 433875.42, '2026-31': 329717.23, '2026-32': 399628.08, '2026-33': 451717.32, '2026-34': 424025.53, '2026-35': 511922.27 },
        inventario: { '2026-28': 700000, '2026-29': 674980.24, '2026-30': 650015.45, '2026-31': 738343.44, '2026-32': 714983.59, '2026-33': 700336.16, '2026-34': 747468.97, '2026-35': 791394.07 }
      },
      J02: {
        recibo: { '2026-28': 165000, '2026-29': 186422, '2026-30': 131837, '2026-31': 159204, '2026-32': 186036, '2026-33': 154109, '2026-34': 159875, '2026-35': 173684 },
        despacho: { '2026-28': 160000, '2026-29': 175196.59, '2026-30': 146341.85, '2026-31': 152808.29, '2026-32': 166084.29, '2026-33': 163976, '2026-34': 154138.53, '2026-35': 172331.42 },
        inventario: { '2026-28': 190000, '2026-29': 196451.31, '2026-30': 180197.75, '2026-31': 183946.36, '2026-32': 193906.26, '2026-33': 191870.59, '2026-34': 189596.8, '2026-35': 177377.93 }
      },
      J05: {
        recibo: { '2026-28': 10000, '2026-29': 13624, '2026-30': 13257, '2026-31': 13158, '2026-32': 7247, '2026-33': 16798, '2026-34': 7876, '2026-35': 22006 },
        despacho: { '2026-28': 12000, '2026-29': 15087, '2026-30': 11608, '2026-31': 9465, '2026-32': 16331, '2026-33': 12065.75, '2026-34': 12478.25, '2026-35': 14712 },
        inventario: { '2026-28': 20000, '2026-29': 21103, '2026-30': 23457, '2026-31': 24903, '2026-32': 16939, '2026-33': 21649, '2026-34': 18557, '2026-35': 23245 }
      },
      J06: {
        recibo: { '2026-28': 6000, '2026-29': 6730, '2026-30': 6913, '2026-31': 5209, '2026-32': 6196, '2026-33': 6249, '2026-34': 6974, '2026-35': 3883 },
        despacho: { '2026-28': 6000, '2026-29': 6527, '2026-30': 6242, '2026-31': 6265, '2026-32': 5813.33, '2026-33': 5391, '2026-34': 6662, '2026-35': 6410 },
        inventario: { '2026-28': 9500, '2026-29': 9780, '2026-30': 9769, '2026-31': 8933, '2026-32': 9935, '2026-33': 9976, '2026-34': 10466, '2026-35': 8499 }
      },
      J07: {
        recibo: { '2026-28': 2000, '2026-29': 2595, '2026-30': 2348, '2026-31': 1460, '2026-32': 2288, '2026-33': 2206, '2026-34': 1316, '2026-35': 2254 },
        despacho: { '2026-28': 2000, '2026-29': 2086, '2026-30': 2499, '2026-31': 1915, '2026-32': 1167, '2026-33': 2178, '2026-34': 2354.6, '2026-35': 2218 },
        inventario: { '2026-28': 2600, '2026-29': 2532, '2026-30': 2555, '2026-31': 2169, '2026-32': 3023, '2026-33': 2869, '2026-34': 2471, '2026-35': 2043 }
      },
      J08: {
        recibo: { '2026-28': 40000, '2026-29': 36632, '2026-30': 62321, '2026-31': 10749, '2026-32': 42211, '2026-33': 64262, '2026-34': 29567, '2026-35': 82616 },
        despacho: { '2026-28': 45000, '2026-29': 51805.53, '2026-30': 92248.34, '2026-31': 34975.5, '2026-32': 43253.5, '2026-33': 61470, '2026-34': 41164.11, '2026-35': 40488.93 },
        inventario: { '2026-28': 180000, '2026-29': 196718, '2026-30': 170742, '2026-31': 152814, '2026-32': 147152.11, '2026-33': 131880.22, '2026-34': 126369, '2026-35': 157498 }
      },
      J09: {
        recibo: { '2026-28': 30000, '2026-29': 28055, '2026-30': 28313, '2026-31': 27038, '2026-32': 29281, '2026-33': 31281, '2026-34': 50307, '2026-35': 37801 },
        despacho: { '2026-28': 24000, '2026-29': 21007.06, '2026-30': 25180.88, '2026-31': 17707.82, '2026-32': 19661.11, '2026-33': 24253.2, '2026-34': 26052.75, '2026-35': 22494.49 },
        inventario: { '2026-28': 120000, '2026-29': 115961.45, '2026-30': 111809.85, '2026-31': 121727.09, '2026-32': 132002.07, '2026-33': 152988.63, '2026-34': 162831.19, '2026-35': 176560.13 }
      },
      J10: {
        recibo: { '2026-28': 15000, '2026-29': 25081, '2026-30': 12448, '2026-31': 3386, '2026-32': 12870, '2026-33': 15297, '2026-34': 5839, '2026-35': 22285 },
        despacho: { '2026-28': 20000, '2026-29': 34890.73, '2026-30': 36185.16, '2026-31': 16453.57, '2026-32': 12222.05, '2026-33': 17214.8, '2026-34': 12023.7, '2026-35': 10443.09 },
        inventario: { '2026-28': 65000, '2026-29': 74289.45, '2026-30': 52869.37, '2026-31': 45005.95, '2026-32': 42676.12, '2026-33': 50206.37, '2026-34': 38911.58, '2026-35': 38872.98 }
      },
      J11: {
        recibo: { '2026-28': 18000, '2026-29': 21781, '2026-30': 17906, '2026-31': 12865, '2026-32': 24117, '2026-33': 18289, '2026-34': 18528, '2026-35': 15563 },
        despacho: { '2026-28': 22000, '2026-29': 33128.39, '2026-30': 30356.4, '2026-31': 18962.06, '2026-32': 25663.55, '2026-33': 24949.74, '2026-34': 18539.71, '2026-35': 16751.39 },
        inventario: { '2026-28': 55000, '2026-29': 64481.63, '2026-30': 51766.76, '2026-31': 46344.06, '2026-32': 44678.1, '2026-33': 42726.55, '2026-34': 43212.37, '2026-35': 44079.94 }
      },
      J12: {
        recibo: { '2026-28': 0, '2026-29': 0, '2026-30': 0, '2026-31': 0, '2026-32': 0, '2026-33': 0, '2026-34': 0, '2026-35': 0 },
        despacho: { '2026-28': 0, '2026-29': 0, '2026-30': 0, '2026-31': 290, '2026-32': 178, '2026-33': 0, '2026-34': 0, '2026-35': 276 },
        inventario: { '2026-28': 98, '2026-29': 98, '2026-30': 98, '2026-31': 98, '2026-32': 76, '2026-33': 98, '2026-34': 98, '2026-35': 98 }
      }
    },
    Frescos: {
      J01: {
        recibo: { '2026-28': 2000, '2026-29': 1996, '2026-30': 1554, '2026-31': 1711, '2026-32': 1531, '2026-33': 2262, '2026-34': 3739, '2026-35': 5369 },
        despacho: { '2026-28': 2200, '2026-29': 2316.88, '2026-30': 1960, '2026-31': 1896, '2026-32': 1662, '2026-33': 1958.75, '2026-34': 2223, '2026-35': 3983 },
        inventario: { '2026-28': 1200, '2026-29': 1414, '2026-30': 1040, '2026-31': 944, '2026-32': 768.75, '2026-33': 706, '2026-34': 902, '2026-35': 4036 }
      },
      J03: {
        recibo: { '2026-28': 22000, '2026-29': 23011.89, '2026-30': 20572.59, '2026-31': 19593, '2026-32': 15055, '2026-33': 24864, '2026-34': 28176, '2026-35': 22424 },
        despacho: { '2026-28': 23000, '2026-29': 21755, '2026-30': 24124, '2026-31': 21794, '2026-32': 24865, '2026-33': 21117.6, '2026-34': 21774, '2026-35': 25010.6 },
        inventario: { '2026-28': 20000, '2026-29': 23110, '2026-30': 18591, '2026-31': 17417, '2026-32': 18121.8, '2026-33': 25108, '2026-34': 25182, '2026-35': 29100 }
      },
      J04: {
        recibo: { '2026-28': 130000, '2026-29': 127521, '2026-30': 130193, '2026-31': 121405, '2026-32': 141980, '2026-33': 140623, '2026-34': 147345, '2026-35': 145552 },
        despacho: { '2026-28': 132000, '2026-29': 131082.52, '2026-30': 131122.9, '2026-31': 119446, '2026-32': 137692.8, '2026-33': 135959.84, '2026-34': 135283.01, '2026-35': 145428.8 },
        inventario: { '2026-28': 13000, '2026-29': 12536.18, '2026-30': 10843, '2026-31': 12151.81, '2026-32': 14899.41, '2026-33': 11196.07, '2026-34': 13652.74, '2026-35': 14228.45 }
      },
      J05: {
        recibo: { '2026-28': 155000, '2026-29': 150602.3, '2026-30': 161279.93, '2026-31': 140408, '2026-32': 148323, '2026-33': 144655, '2026-34': 164669, '2026-35': 192455 },
        despacho: { '2026-28': 150000, '2026-29': 155230.55, '2026-30': 154250.51, '2026-31': 137127.7, '2026-32': 167588.88, '2026-33': 142645.93, '2026-34': 153801, '2026-35': 176621.18 },
        inventario: { '2026-28': 22000, '2026-29': 15843, '2026-30': 26549, '2026-31': 24130, '2026-32': 23030, '2026-33': 17435.2, '2026-34': 21647.2, '2026-35': 26435.2 }
      },
      J06: {
        recibo: { '2026-28': 130000, '2026-29': 139003.88, '2026-30': 131184.18, '2026-31': 136469, '2026-32': 115103, '2026-33': 128154, '2026-34': 133109, '2026-35': 130799 },
        despacho: { '2026-28': 131000, '2026-29': 134868.83, '2026-30': 129342.24, '2026-31': 137464.2, '2026-32': 123301.38, '2026-33': 124484.37, '2026-34': 129218.57, '2026-35': 134799.06 },
        inventario: { '2026-28': 30000, '2026-29': 31399, '2026-30': 32492.5, '2026-31': 29901, '2026-32': 29114.75, '2026-33': 25955, '2026-34': 25501, '2026-35': 31306.45 }
      },
      J07: {
        recibo: { '2026-28': 75000, '2026-29': 81024.68, '2026-30': 88311.9, '2026-31': 86645, '2026-32': 68566, '2026-33': 87569, '2026-34': 75266, '2026-35': 75741 },
        despacho: { '2026-28': 78000, '2026-29': 82939, '2026-30': 82991.58, '2026-31': 91345, '2026-32': 72542, '2026-33': 75344, '2026-34': 75430.75, '2026-35': 80389 },
        inventario: { '2026-28': 20000, '2026-29': 20866, '2026-30': 20003, '2026-31': 20526, '2026-32': 21710, '2026-33': 22425, '2026-34': 20620, '2026-35': 18544 }
      }
    }
  };

  const DEFAULT_SECOS_DATA = {
    headers: ['MES', 'SemAño', 'DESPACHO', 'RECIBO', 'INVENTARIO ACT', 'PLAN RECIBO', 'PLAN DESPACHO', 'PLAN INV'],
    rows: [
      { 'MES': 'Julio', 'SemAño': '[28-2025]', 'DESPACHO': 654590, 'RECIBO': 634039, 'INVENTARIO ACT': 1002142, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Julio', 'SemAño': '[29-2025]', 'DESPACHO': 651786, 'RECIBO': 606364, 'INVENTARIO ACT': 1037701, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Julio', 'SemAño': '[30-2025]', 'DESPACHO': 716068, 'RECIBO': 689254, 'INVENTARIO ACT': 1129480, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Julio', 'SemAño': '[31-2025]', 'DESPACHO': 581563, 'RECIBO': 546992, 'INVENTARIO ACT': 1153051, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Agosto', 'SemAño': '[32-2025]', 'DESPACHO': 688136, 'RECIBO': 616663, 'INVENTARIO ACT': 1119066, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Agosto', 'SemAño': '[33-2025]', 'DESPACHO': 738969, 'RECIBO': 729209, 'INVENTARIO ACT': 1170625, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Agosto', 'SemAño': '[34-2025]', 'DESPACHO': 705780, 'RECIBO': 828357, 'INVENTARIO ACT': 1132918, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Agosto', 'SemAño': '[35-2025]', 'DESPACHO': 751624, 'RECIBO': 834378, 'INVENTARIO ACT': 1120293, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Setiembre', 'SemAño': '[36-2025]', 'DESPACHO': 592920, 'RECIBO': 681043, 'INVENTARIO ACT': 1118982, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Setiembre', 'SemAño': '[37-2025]', 'DESPACHO': 764992, 'RECIBO': 696260, 'INVENTARIO ACT': 1137194, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Setiembre', 'SemAño': '[38-2025]', 'DESPACHO': 789285, 'RECIBO': 931729, 'INVENTARIO ACT': 1263871, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Setiembre', 'SemAño': '[39-2025]', 'DESPACHO': 761992, 'RECIBO': 951931, 'INVENTARIO ACT': 1413284, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      // 2026 Real
      { 'MES': 'Julio', 'SemAño': '[28-2026]', 'DESPACHO': 734961, 'RECIBO': 767742, 'INVENTARIO ACT': 1422509, 'PLAN RECIBO': 770000, 'PLAN DESPACHO': 740000, 'PLAN INV': 1400000 },
      { 'MES': 'Julio', 'SemAño': '[29-2026]', 'DESPACHO': 759228, 'RECIBO': 734703, 'INVENTARIO ACT': 1356395, 'PLAN RECIBO': 750000, 'PLAN DESPACHO': 760000, 'PLAN INV': 1350000 },
      { 'MES': 'Julio', 'SemAño': '[30-2026]', 'DESPACHO': 784537, 'RECIBO': 674119, 'INVENTARIO ACT': 1253280, 'PLAN RECIBO': 720000, 'PLAN DESPACHO': 770000, 'PLAN INV': 1260000 },
      { 'MES': 'Julio', 'SemAño': '[31-2026]', 'DESPACHO': 588559, 'RECIBO': 600920, 'INVENTARIO ACT': 1324284, 'PLAN RECIBO': 620000, 'PLAN DESPACHO': 610000, 'PLAN INV': 1310000 },
      { 'MES': 'Agosto', 'SemAño': '[32-2026]', 'DESPACHO': 690002, 'RECIBO': 702856, 'INVENTARIO ACT': 1305371, 'PLAN RECIBO': 710000, 'PLAN DESPACHO': 700000, 'PLAN INV': 1310000 },
      { 'MES': 'Agosto', 'SemAño': '[33-2026]', 'DESPACHO': 763216, 'RECIBO': 736969, 'INVENTARIO ACT': 1304601, 'PLAN RECIBO': 740000, 'PLAN DESPACHO': 760000, 'PLAN INV': 1300000 },
      { 'MES': 'Agosto', 'SemAño': '[34-2026]', 'DESPACHO': 697439, 'RECIBO': 794188, 'INVENTARIO ACT': 1339982, 'PLAN RECIBO': 800000, 'PLAN DESPACHO': 720000, 'PLAN INV': 1330000 },
      { 'MES': 'Agosto', 'SemAño': '[35-2026]', 'DESPACHO': 798048, 'RECIBO': 887530, 'INVENTARIO ACT': 1419668, 'PLAN RECIBO': 880000, 'PLAN DESPACHO': 790000, 'PLAN INV': 1410000 },
      // 2026 Proyecciones Futuras (Plan)
      { 'MES': 'Setiembre', 'SemAño': '[36-2026]', 'DESPACHO': 0, 'RECIBO': 0, 'INVENTARIO ACT': 0, 'PLAN RECIBO': 820000, 'PLAN DESPACHO': 760000, 'PLAN INV': 1420000 },
      { 'MES': 'Setiembre', 'SemAño': '[37-2026]', 'DESPACHO': 0, 'RECIBO': 0, 'INVENTARIO ACT': 0, 'PLAN RECIBO': 840000, 'PLAN DESPACHO': 780000, 'PLAN INV': 1430000 },
      { 'MES': 'Setiembre', 'SemAño': '[38-2026]', 'DESPACHO': 0, 'RECIBO': 0, 'INVENTARIO ACT': 0, 'PLAN RECIBO': 860000, 'PLAN DESPACHO': 800000, 'PLAN INV': 1440000 },
      { 'MES': 'Setiembre', 'SemAño': '[39-2026]', 'DESPACHO': 0, 'RECIBO': 0, 'INVENTARIO ACT': 0, 'PLAN RECIBO': 850000, 'PLAN DESPACHO': 790000, 'PLAN INV': 1435000 }
    ]
  };

  const DEFAULT_FRESCOS_DATA = {
    headers: ['MES', 'SemAño', 'DESPACHO', 'RECIBO', 'INVENTARIO ACT', 'PLAN RECIBO', 'PLAN DESPACHO', 'PLAN INV'],
    rows: [
      { 'MES': 'Julio', 'SemAño': '[28-2025]', 'DESPACHO': 483972, 'RECIBO': 478512, 'INVENTARIO ACT': 115260, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Julio', 'SemAño': '[29-2025]', 'DESPACHO': 488043, 'RECIBO': 491931, 'INVENTARIO ACT': 115275, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Julio', 'SemAño': '[30-2025]', 'DESPACHO': 467825, 'RECIBO': 490341, 'INVENTARIO ACT': 122200, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Julio', 'SemAño': '[31-2025]', 'DESPACHO': 472321, 'RECIBO': 469514, 'INVENTARIO ACT': 123460, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Agosto', 'SemAño': '[32-2025]', 'DESPACHO': 469799, 'RECIBO': 479278, 'INVENTARIO ACT': 121343, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Agosto', 'SemAño': '[33-2025]', 'DESPACHO': 512125, 'RECIBO': 495154, 'INVENTARIO ACT': 126129, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Agosto', 'SemAño': '[34-2025]', 'DESPACHO': 498158, 'RECIBO': 507833, 'INVENTARIO ACT': 127330, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Agosto', 'SemAño': '[35-2025]', 'DESPACHO': 487991, 'RECIBO': 498404, 'INVENTARIO ACT': 121184, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Setiembre', 'SemAño': '[36-2025]', 'DESPACHO': 439997, 'RECIBO': 472618, 'INVENTARIO ACT': 126689, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Setiembre', 'SemAño': '[37-2025]', 'DESPACHO': 503113, 'RECIBO': 510736, 'INVENTARIO ACT': 121823, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Setiembre', 'SemAño': '[38-2025]', 'DESPACHO': 510860, 'RECIBO': 468129, 'INVENTARIO ACT': 122413, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      { 'MES': 'Setiembre', 'SemAño': '[39-2025]', 'DESPACHO': 463016, 'RECIBO': 500135, 'INVENTARIO ACT': 135537, 'PLAN RECIBO': 0, 'PLAN DESPACHO': 0, 'PLAN INV': 0 },
      // 2026 Real
      { 'MES': 'Julio', 'SemAño': '[28-2026]', 'DESPACHO': 531062, 'RECIBO': 514900, 'INVENTARIO ACT': 111005, 'PLAN RECIBO': 520000, 'PLAN DESPACHO': 530000, 'PLAN INV': 115000 },
      { 'MES': 'Julio', 'SemAño': '[29-2026]', 'DESPACHO': 528193, 'RECIBO': 523160, 'INVENTARIO ACT': 105168, 'PLAN RECIBO': 525000, 'PLAN DESPACHO': 525000, 'PLAN INV': 110000 },
      { 'MES': 'Julio', 'SemAño': '[30-2026]', 'DESPACHO': 523791, 'RECIBO': 533096, 'INVENTARIO ACT': 109519, 'PLAN RECIBO': 530000, 'PLAN DESPACHO': 525000, 'PLAN INV': 112000 },
      { 'MES': 'Julio', 'SemAño': '[31-2026]', 'DESPACHO': 509073, 'RECIBO': 506231, 'INVENTARIO ACT': 105070, 'PLAN RECIBO': 515000, 'PLAN DESPACHO': 515000, 'PLAN INV': 108000 },
      { 'MES': 'Agosto', 'SemAño': '[32-2026]', 'DESPACHO': 527652, 'RECIBO': 490558, 'INVENTARIO ACT': 107645, 'PLAN RECIBO': 500000, 'PLAN DESPACHO': 520000, 'PLAN INV': 110000 },
      { 'MES': 'Agosto', 'SemAño': '[33-2026]', 'DESPACHO': 501510, 'RECIBO': 528127, 'INVENTARIO ACT': 102825, 'PLAN RECIBO': 525000, 'PLAN DESPACHO': 510000, 'PLAN INV': 105000 },
      { 'MES': 'Agosto', 'SemAño': '[34-2026]', 'DESPACHO': 517730, 'RECIBO': 552304, 'INVENTARIO ACT': 107505, 'PLAN RECIBO': 545000, 'PLAN DESPACHO': 525000, 'PLAN INV': 108000 },
      { 'MES': 'Agosto', 'SemAño': '[35-2026]', 'DESPACHO': 566232, 'RECIBO': 572340, 'INVENTARIO ACT': 123650, 'PLAN RECIBO': 565000, 'PLAN DESPACHO': 555000, 'PLAN INV': 120000 },
      // 2026 Proyecciones Futuras (Plan)
      { 'MES': 'Setiembre', 'SemAño': '[36-2026]', 'DESPACHO': 0, 'RECIBO': 0, 'INVENTARIO ACT': 0, 'PLAN RECIBO': 540000, 'PLAN DESPACHO': 530000, 'PLAN INV': 118000 },
      { 'MES': 'Setiembre', 'SemAño': '[37-2026]', 'DESPACHO': 0, 'RECIBO': 0, 'INVENTARIO ACT': 0, 'PLAN RECIBO': 545000, 'PLAN DESPACHO': 535000, 'PLAN INV': 119000 },
      { 'MES': 'Setiembre', 'SemAño': '[38-2026]', 'DESPACHO': 0, 'RECIBO': 0, 'INVENTARIO ACT': 0, 'PLAN RECIBO': 550000, 'PLAN DESPACHO': 540000, 'PLAN INV': 120000 },
      { 'MES': 'Setiembre', 'SemAño': '[39-2026]', 'DESPACHO': 0, 'RECIBO': 0, 'INVENTARIO ACT': 0, 'PLAN RECIBO': 548000, 'PLAN DESPACHO': 538000, 'PLAN INV': 119500 }
    ]
  };

  // ══════════════════════════════════════════════
  // INICIALIZACIÓN DE DATOS Y ESTADO LIMPIO
  // ══════════════════════════════════════════════
  function initThroughputData() {
    // Al cargar la página o refrescar con F5, mantener el dashboard limpio y oculto
    // hasta que el usuario cargue su enlace de Google Sheet o use 'Datos de Prueba'.
    window.dataSecos = null;
    window.dataFrescos = null;
    window.parsedDivisionData = null;
    window.dataDivision = null;

    document.getElementById('dashboardSection')?.classList.add('hidden');
    document.getElementById('connectBox')?.classList.remove('hidden');
    document.getElementById('connectionSuccessInfo')?.classList.add('hidden');
  }

  // Inicializar estado limpio al arrancar
  initThroughputData();
});
