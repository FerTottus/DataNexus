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
    J05: 'FLC (FIAMBRES, LÁCTEOS, CONG.)',
    J06: 'PANADERÍA Y PASTELERÍA',
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

    // 📊 INVENTARIO: Grouped bars — barras agrupadas con espacio suficiente
    renderBarChart(`chart${prefix}Inventario`, labels, [
      { label: `Inventario ${prevYear}`, data: prevInventario, bg: 'rgba(148, 163, 184, 0.65)', border: '#94a3b8' },
      { label: `Inventario ${currentYear}`, data: currInventario, bg: '#059669', border: '#047857' },
    ], planInv.some(v => v > 0) ? { label: 'PLAN INV', data: planInv, color: '#0d9488' } : null, showLabels);

    // 8. Tabla Resumen: El usuario solicitó explícitamente NO incluir la proyección futura en la tabla,
    // sino mostrar únicamente las semanas cerradas reales (específicamente las 8 semanas cerradas).
    const tableWeeks = closedWeeks.length <= 8 ? closedWeeks : closedWeeks.slice(-8);
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

  function getBaseChartOptions(showLabels) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: { top: showLabels ? 48 : 14, bottom: 6, left: 14, right: 14 }
      },
      plugins: {
        legend: {
          position: 'top',
          align: 'center',
          labels: {
            usePointStyle: true,
            boxWidth: 9,
            padding: 18,
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
          display: showLabels,
          clip: false,
          clamp: false
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { font: { weight: '700', size: 11, family: "'Inter', sans-serif" }, color: '#475569' }
        },
        y: {
          beginAtZero: true,
          grace: '14%', // Da margen de respiro superior para que no choque el Plan ni las etiquetas
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
        barPercentage: 0.90,
        categoryPercentage: 0.88,
        datalabels: {
          display: showLabels,
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
          display: showLabels,
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
          display: showLabels,
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

    const isFiltering = selectedDivisions.size > 0 && selectedDivisions.size < 12;
    const sortedCodes = Array.from(selectedDivisions).sort();
    const divTableTag = isFiltering
      ? ` · Filtrando ${sortedCodes.length} ${sortedCodes.length === 1 ? 'División' : 'Divisiones'} (${sortedCodes.join(', ')})`
      : '';

    let html = `
      <div class="table-wrapper-title">
        <i class="fa-solid fa-table-list text-primary"></i>
        <span>Detalle Corporativo Semanal (${currentYear} vs ${prevYear}) · Últimas ${weekNumbers.length} Semanas Cerradas${divTableTag}</span>
      </div>
      <table class="data-table">
        <thead>
          <tr>
            <th style="min-width: 160px;">Métrica / Proceso</th>
    `;
    
    weekNumbers.forEach(w => {
      html += `<th>S${w} (${currentYear})</th>`;
    });
    html += `<th style="min-width:150px; background:#f8fafc;">Promedio Semanal (${weekNumbers.length} Sem.)</th></tr></thead><tbody>`;

    const metrics = [
      { key: 'recibo', planKey: 'planRecibo', label: '📦 RECIBO', color: '#2563eb' },
      { key: 'despacho', planKey: 'planDespacho', label: '🚛 DESPACHO', color: '#ea580c' },
      { key: 'inventario', planKey: 'planInv', label: '📊 INVENTARIO', color: '#059669' }
    ];

    metrics.forEach(m => {
      // 1. Fila Actual (Real) - Promedio Semanal
      html += `<tr><td style="color:${m.color}; font-weight:700;">${m.label} ${currentYear} (Real)</td>`;
      let sumCurr = 0;
      let countCurr = 0;
      weekNumbers.forEach(w => {
        const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][m.key] : 0;
        sumCurr += val;
        if (val > 0) countCurr++;
        html += `<td style="font-weight:600;">${Math.round(val).toLocaleString('es-PE')}</td>`;
      });
      const avgCurr = countCurr > 0 ? sumCurr / countCurr : (weekNumbers.length > 0 ? sumCurr / weekNumbers.length : 0);
      html += `<td style="font-weight:800; background:#f1f5f9;">${Math.round(avgCurr).toLocaleString('es-PE')}</td></tr>`;

      // 2. Fila Plan (Metas del período cerrado) - Promedio Semanal
      html += `<tr style="color:#0284c7; background:rgba(239, 246, 255, 0.35);"><td style="font-weight:600; padding-left: 20px;">└ Plan Objetivo</td>`;
      let sumPlan = 0;
      let countPlan = 0;
      weekNumbers.forEach(w => {
        const val = dataMap[`${currentYear}-${w}`] ? dataMap[`${currentYear}-${w}`][m.planKey] : 0;
        sumPlan += val;
        if (val > 0) countPlan++;
        html += `<td>${val > 0 ? Math.round(val).toLocaleString('es-PE') : '--'}</td>`;
      });
      const avgPlan = countPlan > 0 ? sumPlan / countPlan : 0;
      html += `<td style="font-weight:700; background:#eff6ff;">${avgPlan > 0 ? Math.round(avgPlan).toLocaleString('es-PE') : '--'}</td></tr>`;

      // 3. Fila Año Anterior (Real) - Promedio Semanal
      html += `<tr style="color:#64748b;"><td style="font-weight:600; padding-left: 20px;">└ Real ${prevYear}</td>`;
      let sumPrev = 0;
      let countPrev = 0;
      weekNumbers.forEach(w => {
        const val = dataMap[`${prevYear}-${w}`] ? dataMap[`${prevYear}-${w}`][m.key] : 0;
        sumPrev += val;
        if (val > 0) countPrev++;
        html += `<td>${val > 0 ? Math.round(val).toLocaleString('es-PE') : '--'}</td>`;
      });
      const avgPrev = countPrev > 0 ? sumPrev / countPrev : 0;
      html += `<td style="font-weight:700; background:#f8fafc;">${avgPrev > 0 ? Math.round(avgPrev).toLocaleString('es-PE') : '--'}</td></tr>`;
    });

    html += '</tbody></table>';
    container.innerHTML = html;
  }
});
