/**
 * Script Principal para el Dashboard de Movilidad
 */

// Estado global de la aplicación
const AppState = {
  sheetId: '',
  sheetUrl: '',
  rawEmployees: [], // Todos los empleados combinados
  filteredEmployees: [], // Empleados después de aplicar filtros
  rawRegistroDiario: [], // Filas originales del registro diario
  registroData: [], // Viajes procesados y agregados
  fechasDisponibles: [], // Lista cronológica de fechas únicas
  diasDisponibles: [], // Días de la semana que tienen movimiento real (> 0)
  dateToDiaMap: new Map(), // Mapeo de Fecha normalizada -> Día de la semana
  diaToDatesMap: new Map(), // Mapeo de Día de la semana -> Lista de fechas
  fechaSeleccionada: '', // Fecha activa para el análisis diario
  employeeMap: new Map(), // Mapa rápido de todos los empleados por DNI limpio y original
  charts: {} // Referencias a instancias de Chart.js
};

// ==========================================
// Utilidades de Normalización y Datos
// ==========================================

function cleanDni(val) {
  if (val === undefined || val === null) return '';
  let s = String(val).trim().toUpperCase();
  s = s.replace(/\.0+$/, ''); // Eliminar .0 de números exportados desde Excel
  s = s.replace(/^PE/i, '');  // Quitar prefijo de país Falabella 'PE'
  s = s.replace(/[^0-9A-Z]/g, ''); // Dejar solo caracteres alfanuméricos
  return s;
}

function normalizeTipo(val) {
  if (!val) return 'DESCONOCIDO';
  const s = String(val).trim().toUpperCase();
  if (s.includes('STAFF')) return 'STAFF';
  if (s.includes('OPERAR')) return 'OPERARIO';
  return s;
}

function getRowVal(row, candidates) {
  if (!row) return '';
  const rowKeys = Object.keys(row);
  for (const cand of candidates) {
    const candNorm = cand.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
    for (const k of rowKeys) {
      const kNorm = k.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
      if (kNorm === candNorm && row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
        return row[k];
      }
    }
  }
  return '';
}

function parseDateDMY(dStr) {
  if (!dStr) return 0;
  const s = String(dStr).trim().split(' ')[0];
  // Formato DD/MM/YYYY o DD-MM-YYYY
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dmy) {
    return new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10)).getTime();
  }
  // Formato YYYY/MM/DD o YYYY-MM-DD
  const ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (ymd) {
    return new Date(parseInt(ymd[1], 10), parseInt(ymd[2], 10) - 1, parseInt(ymd[3], 10)).getTime();
  }
  // Número de serie de fecha de Excel (ej: 45455)
  if (/^\d{5}$/.test(s)) {
    return new Date(Math.round((parseInt(s, 10) - 25569) * 86400 * 1000)).getTime();
  }
  const t = Date.parse(s);
  return isNaN(t) ? 0 : t;
}

function normalizeDateStr(val) {
  if (!val) return '';
  const s = String(val).trim().split(' ')[0];
  // Formato DD/MM/YYYY o DD-MM-YYYY
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dmy) {
    const day = dmy[1].padStart(2, '0');
    const month = dmy[2].padStart(2, '0');
    const year = dmy[3];
    return `${day}/${month}/${year}`;
  }
  // Formato YYYY/MM/DD o YYYY-MM-DD
  const ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (ymd) {
    const year = ymd[1];
    const month = ymd[2].padStart(2, '0');
    const day = ymd[3].padStart(2, '0');
    return `${day}/${month}/${year}`;
  }
  // Número de serie de Excel
  if (/^\d{5}$/.test(s)) {
    const date = new Date(Math.round((parseInt(s, 10) - 25569) * 86400 * 1000));
    const day = String(date.getUTCDate()).padStart(2, '0');
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const year = date.getUTCFullYear();
    return `${day}/${month}/${year}`;
  }
  return s;
}

function normalizeDiaStr(dVal, fVal) {
  let s = String(dVal || '').trim().toLowerCase();
  s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (s.startsWith('lun')) return 'Lunes';
  if (s.startsWith('mar')) return 'Martes';
  if (s.startsWith('mie')) return 'Miércoles';
  if (s.startsWith('jue')) return 'Jueves';
  if (s.startsWith('vie')) return 'Viernes';
  if (s.startsWith('sab')) return 'Sábado';
  if (s.startsWith('dom')) return 'Domingo';
  // Si no está el texto del día pero hay fecha, calcularlo
  if (fVal) {
    const ts = parseDateDMY(fVal);
    if (ts) {
      const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
      return dias[new Date(ts).getDay()];
    }
  }
  return '';
}

document.addEventListener('DOMContentLoaded', () => {
  initUIEvents();
  GoogleSheetsService.initAuth(); // Restaura token
  checkAuthAndConfig();
  initCharts();
});

function initUIEvents() {
  document.getElementById('btnGoogleLogin').addEventListener('click', () => {
    GoogleSheetsService.requestAccessToken((success) => {
      if (success) updateAuthUI(true);
    });
  });

  document.getElementById('btnGoogleLogout').addEventListener('click', () => {
    GoogleSheetsService.logout();
    updateAuthUI(false);
  });

  document.getElementById('btnFetchData').addEventListener('click', () => {
    fetchData();
  });

  const btnApply = document.getElementById('btnApplyFilters');
  if(btnApply) {
    btnApply.addEventListener('click', () => {
      applyFilters();
    });
  }

  // Cambio automático y sincronización inteligente de fecha y día
  const selFecha = document.getElementById('filtroFecha');
  if (selFecha) {
    selFecha.addEventListener('change', () => {
      const selDia = document.getElementById('filtroDia');
      const valF = selFecha.value;
      const fechas = AppState.fechasDisponibles || [];
      const targetF = (valF === 'ULTIMA') ? (fechas.length > 0 ? fechas[fechas.length - 1] : '') : valF;

      if (valF !== 'TODAS' && targetF && AppState.dateToDiaMap && selDia) {
        const diaDeFecha = AppState.dateToDiaMap.get(targetF);
        if (diaDeFecha) {
          const optExists = Array.from(selDia.options).some(opt => opt.value === diaDeFecha);
          if (optExists) {
            selDia.value = diaDeFecha;
          }
        }
      }
      applyFilters();
    });
  }

  const selDia = document.getElementById('filtroDia');
  if (selDia) {
    selDia.addEventListener('change', () => {
      const selFecha = document.getElementById('filtroFecha');
      const chosenDia = selDia.value;

      if (chosenDia !== 'TODOS' && AppState.diaToDatesMap && selFecha) {
        const datesForDia = AppState.diaToDatesMap.get(chosenDia) || [];
        const fechas = AppState.fechasDisponibles || [];
        const currentFecha = (selFecha.value === 'ULTIMA')
          ? (fechas.length > 0 ? fechas[fechas.length - 1] : '')
          : selFecha.value;

        // Si la fecha seleccionada actualmente no coincide con el día elegido,
        // cambiamos automáticamente la fecha a la más reciente que sí tiene movimiento ese día
        if (currentFecha !== 'TODAS' && !datesForDia.includes(currentFecha)) {
          if (datesForDia.length > 0) {
            selFecha.value = datesForDia[datesForDia.length - 1];
          }
        }
      }
      applyFilters();
    });
  }

  // Chip Group Logic para multi-selección con actualización inmediata
  document.querySelectorAll('.chip-group').forEach(group => {
    group.addEventListener('click', (e) => {
      if (e.target.classList.contains('chip')) {
        const isTodos = e.target.dataset.value === 'TODOS';
        
        if (isTodos) {
          // Si hace clic en TODOS, deseleccionar los demás
          group.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
          e.target.classList.add('active');
        } else {
          // Si hace clic en otro, deseleccionar TODOS
          const todosChip = group.querySelector('.chip[data-value="TODOS"]');
          if(todosChip) todosChip.classList.remove('active');
          
          // Toggle el botón actual
          e.target.classList.toggle('active');
          
          // Si no queda nada seleccionado, forzar TODOS
          if (group.querySelectorAll('.chip.active').length === 0) {
            if(todosChip) todosChip.classList.add('active');
          }
        }
        applyFilters(); // Aplicar filtros de forma reactiva e inmediata
      }
    });
  });

  const btnToggle = document.getElementById('btnToggleDetalle');
  if(btnToggle) {
    btnToggle.addEventListener('click', () => {
      const container = document.getElementById('contenedorDetalleDistritos');
      if (container.style.display === 'none') {
        container.style.display = 'block';
      } else {
        container.style.display = 'none';
      }
    });
  }

  // Explorar Drive
  document.getElementById('btnBrowseDrive').addEventListener('click', () => {
    if (!GoogleSheetsService.isAuthenticated()) {
      GoogleSheetsService.requestAccessToken((success) => {
        if (success) {
          updateAuthUI(true);
          openDriveModal();
        }
      });
    } else {
      openDriveModal();
    }
  });

  // Modal close
  document.getElementById('close-modal-btn').addEventListener('click', () => {
    document.getElementById('drive-modal').classList.add('hidden');
  });

  document.getElementById('drive-modal').addEventListener('click', (e) => {
    if (e.target.id === 'drive-modal') {
      e.target.classList.add('hidden');
    }
  });

  // Buscador de archivos en modal Drive
  const driveSearchInput = document.getElementById('drive-search-input');
  if (driveSearchInput) {
    driveSearchInput.addEventListener('input', (e) => {
      window.filterDriveList(e.target.value);
    });
  }

  // Toggle Connection Panel
  const toggleConnectionBtn = document.getElementById('toggleConnectionBtn');
  if (toggleConnectionBtn) {
    toggleConnectionBtn.addEventListener('click', () => {
      const body = document.getElementById('connectionCardBody');
      const icon = toggleConnectionBtn.querySelector('i');
      if (body.style.display === 'none') {
        body.style.display = 'block';
        icon.classList.replace('fa-chevron-down', 'fa-chevron-up');
      } else {
        body.style.display = 'none';
        icon.classList.replace('fa-chevron-up', 'fa-chevron-down');
      }
    });
  }

  // Toggle Filters Panel
  const toggleFiltersBtn = document.getElementById('toggleFiltersBtn');
  if (toggleFiltersBtn) {
    toggleFiltersBtn.addEventListener('click', () => {
      const panel = document.getElementById('dashboardFilters');
      const icon = toggleFiltersBtn.querySelector('i');
      if (panel.style.display === 'none') {
        panel.style.display = 'flex';
        icon.classList.replace('fa-chevron-down', 'fa-chevron-up');
      } else {
        panel.style.display = 'none';
        icon.classList.replace('fa-chevron-up', 'fa-chevron-down');
      }
    });
  }
}

async function openDriveModal() {
  const modal = document.getElementById('drive-modal');
  const loading = document.getElementById('drive-loading');
  const empty = document.getElementById('drive-empty');
  const list = document.getElementById('drive-file-list');
  const searchInput = document.getElementById('drive-search-input');
  
  modal.classList.remove('hidden');
  loading.classList.remove('hidden');
  list.innerHTML = '';
  empty.classList.add('hidden');
  if (searchInput) searchInput.value = '';

  const files = await GoogleSheetsService.fetchRecentSpreadsheets();
  
  loading.classList.add('hidden');
  
  if (!files || files.length === 0) {
    empty.classList.remove('hidden');
    return;
  }

  files.forEach(file => {
    const li = document.createElement('li');
    li.className = 'file-list-item';
    li.dataset.filename = file.name.toLowerCase();
    
    const date = new Date(file.modifiedTime).toLocaleDateString('es-ES', { day: 'numeric', month: 'short', year: 'numeric' });
    let isShared = true;
    let ownerName = 'Desconocido';
    
    if (file.owners && file.owners.length > 0) {
      if (file.owners[0].me) {
        isShared = false;
        ownerName = 'Tú';
      } else {
        ownerName = file.owners[0].displayName || 'Compartido';
      }
    }

    const isExcel = file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.name.toLowerCase().endsWith('.xlsx');

    const badgeClass = isShared ? 'badge-shared' : 'badge-owned';
    const badgeText = isShared ? 'Compartido' : 'Mío';
    const iconClass = isExcel ? 'fa-triangle-exclamation text-danger' : (isShared ? 'fa-file-excel file-icon-shared' : 'fa-file-excel file-icon-owned');
    const ownerIcon = isShared ? 'fa-users' : 'fa-user';
    
    const excelWarningHtml = isExcel ? `<span class="file-badge custom-tooltip-icon" style="background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.35); margin-right: 5px;" data-tooltip="Ábrelo en Drive y dale a 'Guardar como Hoja de cálculo de Google'">⚠️ Inválido (.xlsx)</span>` : '';

    li.innerHTML = `
      <div class="file-item-icon">
        <i class="fa-solid ${iconClass}"></i>
      </div>
      <div class="file-item-content">
        <div class="file-item-title" title="${file.name}">${file.name}</div>
        <div class="file-item-meta">
          <span class="owner"><i class="fa-solid ${ownerIcon}"></i> ${ownerName}</span>
          <span class="date"><i class="fa-regular fa-clock"></i> ${date}</span>
        </div>
      </div>
      <div class="file-item-action">
        ${excelWarningHtml}
        <span class="file-badge ${badgeClass}">${badgeText}</span>
      </div>
    `;
    
    li.addEventListener('click', () => {
      document.getElementById('sheetUrlInput').value = file.id;
      modal.classList.add('hidden');
      fetchData(); // Cargar automáticamente
    });
    
    list.appendChild(li);
  });
}

window.filterDriveList = function(searchTerm) {
  const term = String(searchTerm || '')
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  const items = document.querySelectorAll('#drive-file-list .file-list-item');
  let hasVisible = 0;

  items.forEach(item => {
    const rawName = item.dataset.filename || '';
    const normName = rawName.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    if (!term || normName.includes(term)) {
      item.classList.remove('hidden');
      item.style.removeProperty('display');
      hasVisible++;
    } else {
      item.classList.add('hidden');
      item.style.setProperty('display', 'none', 'important');
    }
  });

  const empty = document.getElementById('drive-empty');
  if (empty) {
    if (hasVisible > 0) {
      empty.classList.add('hidden');
      empty.style.display = 'none';
    } else {
      empty.classList.remove('hidden');
      empty.style.display = 'block';
      empty.innerText = term ? `No se encontraron archivos que coincidan con "${searchTerm}"` : 'No se encontraron hojas de cálculo recientes.';
    }
  }
};

function checkAuthAndConfig() {
  const isAuth = GoogleSheetsService.isAuthenticated();
  updateAuthUI(isAuth);
}

function updateAuthUI(isAuthenticated) {
  const btnLogin = document.getElementById('btnGoogleLogin');
  const userInfo = document.getElementById('userInfo');
  const userEmailSpan = document.getElementById('userEmail');

  if (isAuthenticated) {
    btnLogin.classList.add('hidden');
    btnLogin.style.display = 'none';
    userInfo.classList.remove('hidden');
    userInfo.style.display = 'inline-flex';
    userEmailSpan.innerText = GoogleSheetsService.userEmail || 'Conectado';
  } else {
    btnLogin.classList.remove('hidden');
    btnLogin.style.display = '';
    userInfo.classList.add('hidden');
    userInfo.style.display = 'none';
  }
}

async function fetchData() {
  const input = document.getElementById('sheetUrlInput').value.trim();
  if (!input) {
    if(window.ClipboardUtil) ClipboardUtil.showToast('Ingresa la URL del Google Sheet', 'info');
    else alert('Ingresa la URL del Google Sheet');
    return;
  }

  const sheetId = GoogleSheetsService.extractSpreadsheetId(input);
  if (!sheetId) return;

  if (!GoogleSheetsService.isAuthenticated()) {
    GoogleSheetsService.requestAccessToken(async (success) => {
      if (success) {
        updateAuthUI(true);
        await loadAllSheets(sheetId);
      }
    });
    return;
  }

  await loadAllSheets(sheetId);
}

async function loadAllSheets(sheetId) {
  const badge = document.getElementById('connectionStatus');
  badge.className = 'badge badge-warning';
  badge.innerText = 'Descargando datos...';

  try {
    const tabs = await GoogleSheetsService.fetchSheetTabs(sheetId);
    const getTabName = (keywords) => {
      const tab = tabs.find(t => keywords.every(kw => t.title.toUpperCase().includes(kw.toUpperCase())));
      return tab ? tab.title : null;
    };

    const nameSecos = getTabName(['BD SECOS']);
    const namePpa = getTabName(['BD PPA']);
    const nameFrescos = getTabName(['BD FRESCOS']);
    const nameRegistro = getTabName(['REGISTRO', 'DIARIO']); // Busca "REGISTRO_DIARIO", "REGISTRO DIARIO", etc.

    const [secos, ppa, frescos, registroDiario] = await Promise.all([
      nameSecos ? GoogleSheetsService.fetchSheetData(sheetId, nameSecos).catch(e => null) : Promise.resolve(null),
      namePpa ? GoogleSheetsService.fetchSheetData(sheetId, namePpa).catch(e => null) : Promise.resolve(null),
      nameFrescos ? GoogleSheetsService.fetchSheetData(sheetId, nameFrescos).catch(e => null) : Promise.resolve(null),
      nameRegistro ? GoogleSheetsService.fetchSheetData(sheetId, nameRegistro).catch(e => null) : Promise.resolve(null)
    ]);

    if (!secos && !ppa && !frescos) {
      throw new Error("No se pudo leer la información. Si tu archivo es un Excel (.xlsx), debes abrirlo en Google Drive y darle a 'Archivo > Guardar como Hoja de cálculo de Google'. También verifica que existan las pestañas 'BD SECOS', 'BD PPA' o 'BD FRESCOS'.");
    }

    let combined = [];

    const processSheet = (sheetData, areaName) => {
      if (!sheetData || !sheetData.rows) return;
      sheetData.rows.forEach(row => {
        const rawDni = getRowVal(row, ['DNI', 'USERID', 'USER ID', 'DOCUMENTO', 'ID', 'CODIGO']);
        const distrito = String(getRowVal(row, ['Distrito', 'DISTRITO', 'DIST']) || '').trim();
        
        // Ignorar filas completamente vacías o cabeceras repetidas
        if (!distrito && !rawDni) return;
        
        const dniClean = cleanDni(rawDni);
        const tipoRaw = getRowVal(row, ['CLASIFICACION', 'CLASIFICACIÓN', 'TIPO', 'CATEGORIA', 'TIPO EMPLEADO']);
        const tipo = normalizeTipo(tipoRaw);

        const distCdVal = parseFloat(String(getRowVal(row, ['DIST. AL CD (km)', 'DIST. AL CD', 'DIST CD', 'DISTANCIA AL CD']) || '0').replace(/[^0-9.-]+/g, '')) || 0;
        const clasifCd = String(getRowVal(row, ['CLASIF. DIST. CD', 'CLASIF DIST CD', 'CLASIFICACION CD']) || '').trim();
        
        const distParaderoVal = parseFloat(String(getRowVal(row, ['DIST. PARADERO (km)', 'DIST. PARADERO', 'DIST PARADERO', 'DISTANCIA PARADERO']) || '0').replace(/[^0-9.-]+/g, '')) || 0;
        const clasifParadero = String(getRowVal(row, ['CLASIF. DIST. PARADERO', 'CLASIF DIST PARADERO', 'CLASIFICACION PARADERO']) || '').trim();
        
        const ruta = String(getRowVal(row, ['RUTA PARADERO', 'RUTA', 'RUTA ASIGNADA', 'RUTA PREFERIDA']) || '').trim();
        const paradero = String(getRowVal(row, ['PARADERO', 'PARADERO MÁS CERCANO', 'PARADERO MAS CERCANO', 'NOMBRE PARADERO']) || '').trim();

        combined.push({
          dni: dniClean,
          rawDni: String(rawDni).trim(),
          area: areaName,
          tipo: tipo,
          distrito: distrito,
          distCd: distCdVal,
          clasifCd: clasifCd,
          distParadero: distParaderoVal,
          clasifParadero: clasifParadero,
          ruta: ruta,
          paradero: paradero
        });
      });
    };

    processSheet(secos, 'SECOS');
    processSheet(ppa, 'PPA');
    processSheet(frescos, 'FRESCOS');

    AppState.rawEmployees = combined;
    AppState.filteredEmployees = [...combined];
    
    // Construir mapa de empleados para búsqueda rápida por DNI limpio y original
    AppState.employeeMap = new Map();
    combined.forEach(emp => {
      if (emp.dni) AppState.employeeMap.set(emp.dni, emp);
      if (emp.rawDni) AppState.employeeMap.set(String(emp.rawDni).trim().toUpperCase(), emp);
    });

    // Guardamos las filas raw del registro diario para agregarlas dinámicamente según filtros
    AppState.rawRegistroDiario = (registroDiario && registroDiario.rows) ? registroDiario.rows : [];

    // Poblar dropdown de fechas y días basado en los datos únicos normalizados
    const fMap = new Map();
    const diaCounts = {};
    const dateToDia = new Map();
    const diaToDates = new Map();

    AppState.rawRegistroDiario.forEach(r => {
      const rawF = getRowVal(r, ['FECHA', 'FECHA DE VIAJE', 'DATE', 'DIA FECHA']);
      const normF = normalizeDateStr(rawF);
      const rawD = getRowVal(r, ['DÍA', 'DIA', 'DAY']);
      const normD = normalizeDiaStr(rawD, normF);

      if (normF && !fMap.has(normF)) {
        fMap.set(normF, parseDateDMY(normF));
      }

      if (normD) {
        diaCounts[normD] = (diaCounts[normD] || 0) + 1;
        if (normF) {
          dateToDia.set(normF, normD);
          if (!diaToDates.has(normD)) {
            diaToDates.set(normD, []);
          }
          if (!diaToDates.get(normD).includes(normF)) {
            diaToDates.get(normD).push(normF);
          }
        }
      }
    });

    // Orden cronológico estricto (antiguo a reciente)
    const fechas = Array.from(fMap.keys()).sort((a, b) => fMap.get(a) - fMap.get(b));
    AppState.fechasDisponibles = fechas;
    AppState.dateToDiaMap = dateToDia;
    AppState.diaToDatesMap = diaToDates;

    // Asegurar que las fechas dentro de cada día también estén ordenadas cronológicamente
    diaToDates.forEach((datesList) => {
      datesList.sort((a, b) => (fMap.get(a) || 0) - (fMap.get(b) || 0));
    });

    const selFecha = document.getElementById('filtroFecha');
    if (selFecha) {
      const lastDateLabel = fechas.length > 0 ? ` (${fechas[fechas.length - 1]})` : '';
      selFecha.innerHTML = `<option value="ULTIMA">Última Fecha Disponible${lastDateLabel}</option><option value="TODAS">Ver Todas</option>`;
      fechas.forEach(f => {
        selFecha.innerHTML += `<option value="${f}">${f}</option>`;
      });
      selFecha.value = 'ULTIMA';
    }

    // Poblar dropdown de días ÚNICAMENTE con los días que tienen movimiento real (> 0 registros)
    const diasSemanaOrden = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
    const activeDias = diasSemanaOrden.filter(d => (diaCounts[d] || 0) > 0);
    AppState.diasDisponibles = activeDias;

    const selDia = document.getElementById('filtroDia');
    if (selDia) {
      selDia.innerHTML = `<option value="TODOS">Todos los Días</option>`;
      activeDias.forEach(d => {
        selDia.innerHTML += `<option value="${d}">${d}</option>`;
      });
      selDia.value = 'TODOS';
    }

    badge.className = 'badge badge-success';
    badge.innerHTML = `<i class="fa-solid fa-check"></i> Conectado - ${combined.length} regs`;

    document.getElementById('dashboardFiltersWrapper').classList.remove('hidden');
    document.getElementById('dashboardContent').classList.remove('hidden');

    applyFilters();

    if(window.ClipboardUtil) ClipboardUtil.showToast('Datos cargados exitosamente', 'success');

  } catch (err) {
    console.error(err);
    badge.className = 'badge badge-danger';
    badge.innerText = 'Error de conexión';
    if(window.ClipboardUtil) {
      ClipboardUtil.showToast(err.message, 'danger', 10000);
    } else {
      alert(err.message);
    }
  }
}

function applyFilters() {
  const getActiveChips = (containerId) => {
    const activeBtn = document.querySelectorAll(`#${containerId} .chip.active`);
    return Array.from(activeBtn).map(b => b.dataset.value);
  };

  const areas = getActiveChips('chipArea');
  const tipos = getActiveChips('chipTipo');

  const filterAreaActive = !areas.includes('TODOS') && areas.length > 0;
  const filterTipoActive = !tipos.includes('TODOS') && tipos.length > 0;
  const isFiltered = filterAreaActive || filterTipoActive;

  // Filtrar data demográfica de empleados
  AppState.filteredEmployees = AppState.rawEmployees.filter(emp => {
    const matchArea = !filterAreaActive || areas.includes(emp.area);
    const matchTipo = !filterTipoActive || tipos.includes(emp.tipo);
    return matchArea && matchTipo;
  });

  // Mapa de DNIs de empleados filtrados para consulta rápida
  const filteredDniSet = new Set();
  AppState.filteredEmployees.forEach(e => {
    if (e.dni) filteredDniSet.add(e.dni);
    if (e.rawDni) filteredDniSet.add(String(e.rawDni).trim().toUpperCase());
  });

  const selF = document.getElementById('filtroFecha');
  const valFecha = selF ? selF.value : 'ULTIMA';
  const selD = document.getElementById('filtroDia');
  const valDia = selD ? selD.value : 'TODOS';

  const fechas = AppState.fechasDisponibles || [];
  const ultimaFecha = fechas.length > 0 ? fechas[fechas.length - 1] : '';
  const fechaTarget = (valFecha === 'ULTIMA') ? ultimaFecha : valFecha;
  AppState.fechaSeleccionada = fechaTarget;

  // Agrupación de viajes de buses
  // Cada viaje de bus se identifica por fecha + ruta (+ tipoBus si aplica)
  const aggrMap = {};
  AppState.rawRegistroDiario.forEach(row => {
    const rawFecha = getRowVal(row, ['FECHA', 'FECHA DE VIAJE', 'DATE', 'DIA FECHA']);
    const fechaSoloDia = normalizeDateStr(rawFecha);
    if (!fechaSoloDia) return;

    const rawDia = getRowVal(row, ['DÍA', 'DIA', 'DAY']);
    const diaValNorm = normalizeDiaStr(rawDia, fechaSoloDia);

    // Filtros de fecha y día
    if (valDia !== 'TODOS') {
      const valDiaNorm = normalizeDiaStr(valDia);
      if (diaValNorm !== valDiaNorm) return;
    }
    if (fechaTarget !== 'TODAS' && fechaTarget && fechaSoloDia !== fechaTarget) return;

    const rutaVal = String(getRowVal(row, ['RUTA', 'RUTA ASIGNADA', 'LINEA']) || '').trim();
    if (!rutaVal) return;

    const tipoBusVal = String(getRowVal(row, ['TIPO_BUS', 'TIPO BUS', 'BUS_TIPO']) || '').trim();
    const tripKey = `${fechaSoloDia}|${rutaVal}|${tipoBusVal || 'BUS'}`;

    if (!aggrMap[tripKey]) {
      const rawCosto = getRowVal(row, ['COSTO TOTAL', 'COSTO', 'COSTO POR VIAJE', 'COSTO BUS', 'COSTO_TOTAL', 'COSTO IDA Y VUELTA']);
      const costoNum = parseFloat(String(rawCosto || '0').replace(/[^0-9.-]+/g, "")) || 0;
      const rawCap = getRowVal(row, ['CAPACIDAD', 'CAPACIDAD DE BUS', 'CAPACIDAD BUS', 'CAPACIDAD_BUS']);
      const capNum = parseFloat(String(rawCap || '0').replace(/[^0-9.-]+/g, "")) || 0;
      const semVal = parseInt(getRowVal(row, ['SEMANA', 'SEM'])) || 0;

      aggrMap[tripKey] = {
        dia: diaValNorm || rawDia || '',
        fecha: fechaSoloDia,
        semana: semVal,
        ruta: rutaVal,
        tipoBus: tipoBusVal,
        capacidad: capNum,
        costoBus: costoNum,
        totalPasajeros: 0,
        pasajerosFiltrados: 0
      };
    }

    const trip = aggrMap[tripKey];
    trip.totalPasajeros += 1;

    // Verificar si el empleado de esta fila cumple con los filtros activos
    const rawRowDni = getRowVal(row, ['DNI', 'USERID', 'USER ID', 'DOCUMENTO', 'ID', 'CODIGO']);
    const cleanPassengerDni = cleanDni(rawRowDni);

    let passengerMatches = false;
    if (!isFiltered) {
      passengerMatches = true;
    } else {
      if (cleanPassengerDni && filteredDniSet.has(cleanPassengerDni)) {
        passengerMatches = true;
      } else if (rawRowDni && filteredDniSet.has(String(rawRowDni).trim().toUpperCase())) {
        passengerMatches = true;
      }
    }

    if (passengerMatches) {
      trip.pasajerosFiltrados += 1;
    }
  });

  // Crear la data final de viajes
  AppState.registroData = Object.values(aggrMap).map(trip => {
    const prop = trip.totalPasajeros > 0 ? (trip.pasajerosFiltrados / trip.totalPasajeros) : 0;
    // Si hay filtros, prorrateamos el costo según la cantidad de pasajeros que corresponden a ese grupo
    const costoFinal = isFiltered ? (trip.costoBus * prop) : trip.costoBus;
    const pasajerosFinal = isFiltered ? trip.pasajerosFiltrados : trip.totalPasajeros;

    return {
      dia: trip.dia,
      fecha: trip.fecha,
      semana: trip.semana,
      ruta: trip.ruta,
      tipoBus: trip.tipoBus,
      capacidad: trip.capacidad,
      totalPasajerosBus: trip.totalPasajeros,
      pasajeros: pasajerosFinal,
      costo: costoFinal,
      costoBusOriginal: trip.costoBus,
      esFiltrado: isFiltered
    };
  });

  renderTables();
  if (typeof renderCharts === 'function') {
    renderCharts();
  }
}

function renderTables() {
  const data = AppState.filteredEmployees;
  const regData = AppState.registroData;

  const totalEmps = data.length;

  let sumDistCd = 0, countValidCd = 0;
  let sumDistParadero = 0, countValidParadero = 0;
  let distritosSet = new Set();
  let rutasSet = new Set();

  // Variables para tablas de clasificación
  const countCd = { '🟢 Muy Cerca': 0, '🟡 Cerca': 0, '🟠 Moderada': 0, '🔴 Lejos': 0 };
  const countParadero = { '🟢 Muy Cerca': 0, '🟡 Cerca': 0, '🟠 Moderada': 0, '🔴 Lejos': 0 };

  // Variables para tablas detalladas
  const distritosStats = {};
  const rutasCount = {};
  const paraderosCount = {};

  data.forEach(emp => {
    // Totales
    if (emp.distCd > 0) { sumDistCd += emp.distCd; countValidCd++; }
    if (emp.distParadero > 0) { sumDistParadero += emp.distParadero; countValidParadero++; }
    if (emp.distrito) distritosSet.add(emp.distrito);
    if (emp.ruta) rutasSet.add(emp.ruta);

    // Clasificaciones (matching substring to ensure we map correctly, as raw data might just be "Muy Cerca")
    if (emp.clasifCd.includes('Muy Cerca')) countCd['🟢 Muy Cerca']++;
    else if (emp.clasifCd.includes('Cerca')) countCd['🟡 Cerca']++;
    else if (emp.clasifCd.includes('Moderada')) countCd['🟠 Moderada']++;
    else if (emp.clasifCd.includes('Lejos')) countCd['🔴 Lejos']++;

    if (emp.clasifParadero.includes('Muy Cerca')) countParadero['🟢 Muy Cerca']++;
    else if (emp.clasifParadero.includes('Cerca')) countParadero['🟡 Cerca']++;
    else if (emp.clasifParadero.includes('Moderada')) countParadero['🟠 Moderada']++;
    else if (emp.clasifParadero.includes('Lejos')) countParadero['🔴 Lejos']++;

    // Stats por Distrito
    if (emp.distrito) {
      if (!distritosStats[emp.distrito]) {
        distritosStats[emp.distrito] = {
          count: 0, sumCd: 0, validCd: 0, sumPar: 0, validPar: 0, rutas: {}
        };
      }
      const dStat = distritosStats[emp.distrito];
      dStat.count++;
      if (emp.distCd > 0) { dStat.sumCd += emp.distCd; dStat.validCd++; }
      if (emp.distParadero > 0) { dStat.sumPar += emp.distParadero; dStat.validPar++; }
      if (emp.ruta) { dStat.rutas[emp.ruta] = (dStat.rutas[emp.ruta] || 0) + 1; }
    }

    // Listas simples
    if (emp.ruta) rutasCount[emp.ruta] = (rutasCount[emp.ruta] || 0) + 1;
    if (emp.paradero) {
      if (!paraderosCount[emp.paradero]) paraderosCount[emp.paradero] = { count: 0, ruta: emp.ruta || '-' };
      paraderosCount[emp.paradero].count++;
    }
  });

  // 1. Tabla RESUMEN (Ahora KPIs Superiores)
  const elResTotalEmp = document.getElementById('kpiTotalEmps');
  if(elResTotalEmp) {
    elResTotalEmp.innerText = totalEmps;
    document.getElementById('kpiDistPromCD').innerText = countValidCd > 0 ? (sumDistCd / countValidCd).toFixed(2) : '0';
  }

  // Helpers de Formato
  const formatPct = (val, tot) => tot > 0 ? ((val / tot) * 100).toFixed(2) + '%' : '0%';
  const clasificarCD = (km) => {
    if (km <= 5) return '🟢 Muy Cerca';
    if (km <= 10) return '🟡 Cerca';
    if (km <= 20) return '🟠 Moderada';
    return '🔴 Lejos';
  };
  const clasificarPar = (km) => {
    if (km <= 1) return '🟢 Muy Cerca';
    if (km <= 3) return '🟡 Cerca';
    if (km <= 5) return '🟠 Moderada';
    return '🔴 Lejos';
  };

  const obsDistrito = (cd, par) => {
    // Lógica original de Excel provista por el usuario
    if (cd > 20 && par > 5) return '⚠️ Muy lejos del CD y paraderos';
    if (cd > 20) return '⚠️ Muy lejos del CD';
    if (par > 5) return '⚠️ Lejos de paraderos';
    return '📍 Normal';
  };

  // 2 y 3. Gráficos de Clasificación
  const colors = {
    '🟢 Muy Cerca': '#22c55e', 
    '🟡 Cerca': '#eab308',     
    '🟠 Moderada': '#f97316',  
    '🔴 Lejos': '#ef4444'      
  };
  renderPieChart('chartDistCD', countCd, colors);
  renderPieChart('chartDistParadero', countParadero, colors);

  // 4. Tabla Análisis Detallado por Distrito
  const distritosArr = Object.keys(distritosStats).map(d => {
    const st = distritosStats[d];
    const promCd = st.validCd > 0 ? st.sumCd / st.validCd : 0;
    const promPar = st.validPar > 0 ? st.sumPar / st.validPar : 0;
    
    // ruta principal
    let mainRuta = '-';
    let maxRutaCount = 0;
    for (const [r, c] of Object.entries(st.rutas)) {
      if (c > maxRutaCount) { maxRutaCount = c; mainRuta = r; }
    }

    return {
      nombre: d,
      count: st.count,
      promCd: promCd,
      promPar: promPar,
      ruta: mainRuta
    };
  }).sort((a, b) => b.count - a.count); // Ordenar por volumen

  const elTbDetalle = document.querySelector('#tableDetalleDistritos tbody');
  if(elTbDetalle) {
    elTbDetalle.innerHTML = distritosArr.map((d, i) => `
      <tr>
        <td>${i + 1}</td>
        <td>${d.nombre}</td>
        <td>${d.count}</td>
        <td>${formatPct(d.count, totalEmps)}</td>
        <td>${d.promCd >= 0 ? d.promCd.toFixed(2) : '-'}</td>
        <td>${d.promCd >= 0 ? clasificarCD(d.promCd) : '-'}</td>
        <td>${d.promPar >= 0 ? d.promPar.toFixed(2) : '-'}</td>
        <td>${d.promPar >= 0 ? clasificarPar(d.promPar) : '-'}</td>
        <td>${d.ruta}</td>
        <td>${obsDistrito(d.promCd, d.promPar)}</td>
      </tr>
    `).join('');
  }

  // 5. TOP 10 Distritos (re-uso de distritosArr)
  const top10Distritos = distritosArr.slice(0, 10);
  renderBarChart('chartTopDistritos', top10Distritos);

  // 6. Rutas Disponibles
  const rutasSorted = Object.keys(rutasCount).map(k => ({ nombre: k, count: rutasCount[k] })).sort((a, b) => b.count - a.count);
  const elTbRutas = document.querySelector('#tableRutasDisponibles tbody');
  if(elTbRutas) {
    elTbRutas.innerHTML = rutasSorted.map((r, i) => `
      <tr><td>${r.nombre}</td><td>${r.count}</td><td>${formatPct(r.count, totalEmps)}</td></tr>
    `).join('');
  }

  // 7. Top 10 Paraderos
  const paraderosSorted = Object.keys(paraderosCount).map(k => ({ nombre: k, count: paraderosCount[k].count, ruta: paraderosCount[k].ruta })).sort((a, b) => b.count - a.count).slice(0, 10);
  const elTbParaderosTop = document.querySelector('#tableTopParaderos tbody');
  if(elTbParaderosTop) {
    elTbParaderosTop.innerHTML = paraderosSorted.map((p, i) => `
      <tr><td>${i + 1}</td><td>${p.nombre}</td><td>${p.count}</td><td>${p.ruta}</td></tr>
    `).join('');
  }

  // 8. Dashboard Registro Diario
  const seccionReg = document.getElementById('seccionRegistroDiario');
  const seccionBuses = document.getElementById('seccionDashboardBuses');

  if (AppState.rawRegistroDiario && AppState.rawRegistroDiario.length > 0) {
    if (seccionReg) seccionReg.style.display = 'block';
    if (seccionBuses) seccionBuses.style.display = 'block';

    // Fecha a mostrar en el resumen del día
    let fechaActual = AppState.fechaSeleccionada;
    const isVerTodas = (fechaActual === 'TODAS');
    if (!fechaActual || isVerTodas) {
      const fechasUnicas = Array.from(new Set(regData.map(r => r.fecha))).sort((a, b) => parseDateDMY(a) - parseDateDMY(b));
      fechaActual = isVerTodas ? 'TODAS LAS FECHAS' : (fechasUnicas.length > 0 ? fechasUnicas[fechasUnicas.length - 1] : '');
    }

    const isFiltered = regData.some(r => r.esFiltrado);

    // Filtrar viajes del día seleccionado (o todos si se seleccionó TODAS)
    const viajesDia = isVerTodas ? regData : regData.filter(r => r.fecha === fechaActual);
    const datosDia = isFiltered ? viajesDia.filter(r => r.pasajeros > 0) : viajesDia;

    // Obtener la semana del día seleccionado (o última semana registrada)
    const semanaActual = viajesDia.length > 0 ? viajesDia[viajesDia.length - 1].semana : (regData.length > 0 ? regData[regData.length - 1].semana : 0);
    const viajesSemana = isVerTodas ? regData : regData.filter(r => r.semana === semanaActual);
    const datosSemana = isFiltered ? viajesSemana.filter(r => r.pasajeros > 0) : viajesSemana;

    // Actualizar encabezados con la fecha y semana seleccionada
    const elTituloReg = document.querySelector('#seccionRegistroDiario .table-title');
    if (elTituloReg) {
      elTituloReg.innerText = `EFICIENCIA Y COSTO POR RUTA (${isVerTodas ? 'TODAS LAS FECHAS' : 'DÍA: ' + (fechaActual || 'ACTUAL')})`;
    }
    const elResDiaTable = document.getElementById('tableResumenDia');
    if (elResDiaTable) {
      const h4 = elResDiaTable.closest('div').querySelector('h4');
      if (h4) h4.innerHTML = `<i class="fa-regular fa-calendar-days"></i> ${isVerTodas ? 'RESUMEN ACUMULADO (TODAS LAS FECHAS)' : 'RESUMEN DEL DÍA (' + (fechaActual || 'ACTUAL') + ')'}`;
    }
    const elResSemTable = document.getElementById('tableResumenSemana');
    if (elResSemTable) {
      const h4 = elResSemTable.closest('div').querySelector('h4');
      if (h4) h4.innerHTML = `<i class="fa-solid fa-calendar-week"></i> ${isVerTodas ? 'RESUMEN GENERAL' : 'RESUMEN SEMANAL (SEMANA ' + (semanaActual || '-') + ')'}`;
    }

    // Resumen por ruta del DÍA actual
    const statsRuta = {};
    datosDia.forEach(r => {
      if (!statsRuta[r.ruta]) {
        statsRuta[r.ruta] = { viajes: 0, cap: 0, pasaj: 0, costo: 0, pMin: 9999, pMax: -1 };
      }
      const st = statsRuta[r.ruta];
      st.viajes++;
      st.cap += r.capacidad;
      st.pasaj += r.pasajeros;
      st.costo += r.costo;
      if (r.pasajeros < st.pMin) st.pMin = r.pasajeros;
      if (r.pasajeros > st.pMax) st.pMax = r.pasajeros;
    });

    const rutasDiario = Object.keys(statsRuta).map(k => ({ ruta: k, ...statsRuta[k] })).sort((a, b) => a.ruta.localeCompare(b.ruta));
    
    const elTbReg = document.querySelector('#tableRegistroDiario tbody');
    if(elTbReg) {
      if (rutasDiario.length === 0) {
        elTbReg.innerHTML = `<tr><td colspan="10" style="text-align: center; color: #94a3b8; padding: 20px;">No se registraron viajes para los filtros seleccionados en esta fecha (${fechaActual}).</td></tr>`;
      } else {
        elTbReg.innerHTML = rutasDiario.map(r => {
          const pOcup = r.cap > 0 ? r.pasaj / r.cap : 0;
          const cProm = r.pasaj > 0 ? r.costo / r.pasaj : 0;
          const promPasaj = r.viajes > 0 ? r.pasaj / r.viajes : 0;
          const cViaje = r.viajes > 0 ? r.costo / r.viajes : 0;
          
          let estCap = '-';
          if (pOcup >= 0.9) estCap = '✅ Óptimo';
          else if (pOcup >= 0.7) estCap = '🔄 OK';
          else if (pOcup >= 0.5) estCap = '⚠️ Bajo';
          else estCap = '🔴 Muy Bajo';

          return `<tr>
            <td>${r.ruta}</td>
            <td>${r.viajes}</td>
            <td>${r.cap}</td>
            <td>${r.pasaj}</td>
            <td>${(pOcup*100).toFixed(2)}%</td>
            <td>${estCap}</td>
            <td>S/ ${r.costo.toFixed(2)}</td>
            <td>S/ ${cProm.toFixed(2)}</td>
            <td>S/ ${cViaje.toFixed(2)}</td>
            <td>MAX: ${r.pMax === -1 ? 0 : r.pMax} - MIN: ${r.pMin === 9999 ? 0 : r.pMin}</td>
          </tr>`;
        }).join('');
      }
    }

    // Dashboard Día / Semana
    const aggDia = { viajes: 0, pasaj: 0, costo: 0, cap: 0 };
    datosDia.forEach(r => { aggDia.viajes++; aggDia.pasaj += r.pasajeros; aggDia.costo += r.costo; aggDia.cap += r.capacidad; });
    
    const aggSem = { viajes: 0, pasaj: 0, costo: 0, cap: 0 };
    datosSemana.forEach(r => { aggSem.viajes++; aggSem.pasaj += r.pasajeros; aggSem.costo += r.costo; aggSem.cap += r.capacidad; });

    // Actualizar KPIs de la parte superior
    document.getElementById('kpiCostoDia').innerText = 'S/ ' + aggDia.costo.toFixed(2);
    document.getElementById('kpiOcupacionDia').innerText = formatPct(aggDia.pasaj, aggDia.cap);
    const subCosto = document.querySelector('#kpiCostoDia + .kpi-subtitle');
    if (subCosto) subCosto.innerText = isVerTodas ? 'Total invertido acumulado' : 'Total invertido hoy';
    const subOcup = document.querySelector('#kpiOcupacionDia + .kpi-subtitle');
    if (subOcup) subOcup.innerText = isVerTodas ? 'Eficiencia acumulada' : 'Eficiencia operativa hoy';

    // Tablas de Dashboard
    const elRdBusesDia = document.getElementById('rdBusesDia');
    if(elRdBusesDia) {
      elRdBusesDia.innerText = aggDia.viajes;
      document.getElementById('rdPasajerosDia').innerText = aggDia.pasaj;
      document.getElementById('rdCapacidadDia').innerText = aggDia.cap;
      document.getElementById('rdCostoPasajeroDia').innerText = 'S/ ' + (aggDia.pasaj > 0 ? (aggDia.costo / aggDia.pasaj).toFixed(2) : '0.00');
      document.getElementById('rdCapacidadUsadaDia').innerText = `${aggDia.pasaj} de ${aggDia.cap} (${formatPct(aggDia.pasaj, aggDia.cap)})`;

      document.getElementById('rdViajesSem').innerText = aggSem.viajes;
      document.getElementById('rdPasajerosSem').innerText = aggSem.pasaj;
      document.getElementById('rdCostoSem').innerText = 'S/ ' + aggSem.costo.toFixed(2);
      document.getElementById('rdOcupacionSem').innerText = formatPct(aggSem.pasaj, aggSem.cap);
      document.getElementById('rdCostoPasajeroSem').innerText = 'S/ ' + (aggSem.pasaj > 0 ? (aggSem.costo / aggSem.pasaj).toFixed(2) : '0.00');
    }
  } else {
    if(seccionReg) seccionReg.style.display = 'none';
    if(seccionBuses) seccionBuses.style.display = 'none';
  }
}

function initCharts() {
  Chart.defaults.font.family = "'Inter', sans-serif";
  Chart.defaults.color = '#94a3b8';
  if (Chart.defaults.scale && Chart.defaults.scale.grid) {
    Chart.defaults.scale.grid.color = 'rgba(255, 255, 255, 0.08)';
  }
}

function renderPieChart(canvasId, dataMap, colors) {
  const ctx = document.getElementById(canvasId);
  if(!ctx) return;
  
  const labels = Object.keys(dataMap);
  const data = labels.map(k => dataMap[k]);
  const bgColors = labels.map(k => colors[k] || '#cbd5e1');

  if (AppState.charts[canvasId]) {
    AppState.charts[canvasId].destroy();
  }

  AppState.charts[canvasId] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: bgColors,
        borderColor: '#1e293b',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: {
          top: 22,
          bottom: 12,
          left: 10,
          right: 10
        }
      },
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: '#cbd5e1',
            boxWidth: 14,
            padding: 18,
            font: {
              size: 12,
              weight: '500'
            }
          }
        }
      }
    }
  });
}

function renderBarChart(canvasId, dataArr) {
  const ctx = document.getElementById(canvasId);
  if(!ctx) return;
  
  const labels = dataArr.map(d => d.nombre);
  const data = dataArr.map(d => d.count);

  if (AppState.charts[canvasId]) {
    AppState.charts[canvasId].destroy();
  }

  AppState.charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Volumen de Empleados',
        data: data,
        backgroundColor: '#3b82f6',
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0 } }
      },
      plugins: {
        legend: { display: false }
      }
    }
  });
}

