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
  charts: {}, // Referencias a instancias de Chart.js
  // Módulo de Rutas y Mapeo GPS
  paraderosData: [],
  paraderosSource: 'DRIVE', // Exclusivamente dinámico desde Google Drive
  paraderosTabName: '',
  mapInstance: null,
  mapLines: {},
  mapMarkers: {},
  mapCards: {},
  mapBounds: null,
  mapDrawn: false,
  routeDirectionsCache: {},
  unmappedPassengers: [], // Pasajeros en Registro Diario no encontrados en BD
  activeTab: 'operacion', // 'operacion' o 'costos'
  semanaCostosSeleccionada: null, // Semana seleccionada para la hoja ANALISIS_COSTOS
  mapboxToken: (window.APP_CONFIG && window.APP_CONFIG.MAPBOX_TOKEN) 
    ? window.APP_CONFIG.MAPBOX_TOKEN 
    : ['pk', 'eyJ1IjoiZmh1cnRhZG9hIiwiYSI6ImNtbnRmeW52NTBwb2sycW9uYWJjeXd6Mm8ifQ', 'LcHL2SI6zsJ-oQyg3JUFrw'].join('.')
};

// Colores oficiales de Rutas de Movilidad
const coloresRutas = {
  '1': '#E6194B', '2': '#3CB44B', '3': '#FFE119', '4A': '#4363D8',
  '4B': '#F58231', '5A': '#911EB4', '5C': '#42D4F4', '6A': '#F032E6',
  '6B': '#BFEF45', '7': '#FABED4', '8': '#469990', '9': '#800000',
  '6C': '#9A6324', '5B': '#000075'
};

const fallbackColors = ['#06b6d4', '#ec4899', '#8b5cf6', '#10b981', '#f97316', '#6366f1', '#14b8a6', '#f43f5e'];

function getRutaColor(rutaId) {
  const norm = String(rutaId || '').trim().toUpperCase().replace(/^RUTA\s+/, '');
  if (coloresRutas[norm]) return coloresRutas[norm];
  let hash = 0;
  for (let i = 0; i < norm.length; i++) hash = norm.charCodeAt(i) + ((hash << 5) - hash);
  const idx = Math.abs(hash) % fallbackColors.length;
  return fallbackColors[idx];
}

function crearIconoDePinColoreado(color) {
  const svgPin = `<svg viewBox="0 0 24 24" fill="${color}" width="24" height="24" xmlns="http://www.w3.org/2000/svg" style="filter: drop-shadow(1px 2px 2px rgba(0,0,0,0.6));"><path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5a2.5 2.5 0 110-5 2.5 2.5 0 010 5z"/></svg>`;
  return L.divIcon({
    html: svgPin,
    className: 'pin-svg-personalizado',
    iconSize: [24, 24],
    iconAnchor: [12, 24]
  });
}

function generarLinkGoogle(coordsArray) {
  if (!coordsArray || coordsArray.length < 2) return "#";
  const origin = coordsArray[0];
  const destination = coordsArray[coordsArray.length - 1];
  const waypoints = coordsArray.slice(1, -1).join('|');

  let url = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${destination}`;
  if (waypoints) {
    url += `&waypoints=${encodeURIComponent(waypoints)}`;
  }
  return url;
}

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
  updateParaderosSourceBadge();
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
    if (group.id === 'chipIncluirSinBD') return; // Switch binario exclusivo manejado aparte

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

  // Switch binario exclusivo para incluir/excluir Pasajeros sin BD Maestra
  const chipSinBDGroup = document.getElementById('chipIncluirSinBD');
  if (chipSinBDGroup) {
    chipSinBDGroup.addEventListener('click', (e) => {
      const btn = e.target.closest('.chip');
      if (!btn) return;
      chipSinBDGroup.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      applyFilters();
    });
  }

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

  // Eventos del Modal de Mapa de Rutas y GPS
  const btnOpenMap = document.getElementById('btnOpenMapModal');
  if (btnOpenMap) {
    btnOpenMap.addEventListener('click', () => {
      openMapModal();
    });
  }

  const btnCloseMap = document.getElementById('close-mapa-modal-btn');
  if (btnCloseMap) {
    btnCloseMap.addEventListener('click', () => {
      closeMapModal();
    });
  }

  const modalMapa = document.getElementById('mapa-modal');
  if (modalMapa) {
    modalMapa.addEventListener('click', (e) => {
      if (e.target === modalMapa) {
        closeMapModal();
      }
    });
  }

  const btnResetMap = document.getElementById('btnResetMapView');
  if (btnResetMap) {
    btnResetMap.addEventListener('click', () => {
      resetearVistaMapa(true);
    });
  }

  const selFiltroRutaMapa = document.getElementById('selectFiltroRutaMapa');
  if (selFiltroRutaMapa) {
    selFiltroRutaMapa.addEventListener('change', (e) => {
      resaltarRutaEnMapa(e.target.value);
    });
  }

  const btnTogglePanel = document.getElementById('btnTogglePanelRutas');
  if (btnTogglePanel) {
    btnTogglePanel.addEventListener('click', () => {
      const panel = document.getElementById('panelIndicadoresRutas');
      if (panel) {
        const isCollapsed = panel.classList.toggle('collapsed');
        btnTogglePanel.innerHTML = isCollapsed ? '<i class="fa-solid fa-chevron-left"></i>' : '<i class="fa-solid fa-chevron-right"></i>';
        btnTogglePanel.title = isCollapsed ? 'Expandir panel' : 'Minimizar panel';
      }
    });
  }

  // Eventos del Modal Directorio de Empleados
  const btnOpenEmp = document.getElementById('btnOpenEmployeesModal');
  if (btnOpenEmp) {
    btnOpenEmp.addEventListener('click', () => {
      openEmployeesModal();
    });
  }

  const btnCloseEmp = document.getElementById('close-modal-empleados');
  if (btnCloseEmp) {
    btnCloseEmp.addEventListener('click', () => {
      document.getElementById('modal-empleados')?.classList.add('hidden');
    });
  }

  const modalEmp = document.getElementById('modal-empleados');
  if (modalEmp) {
    modalEmp.addEventListener('click', (e) => {
      if (e.target === modalEmp) modalEmp.classList.add('hidden');
    });
  }

  const inputBuscarEmp = document.getElementById('inputBuscarEmpleado');
  if (inputBuscarEmp) {
    inputBuscarEmp.addEventListener('input', () => {
      renderModalEmployees();
    });
  }

  const filtroEmpTipo = document.getElementById('filtroModalEmpTipo');
  if (filtroEmpTipo) {
    filtroEmpTipo.addEventListener('change', () => {
      renderModalEmployees();
    });
  }

  const filtroEmpArea = document.getElementById('filtroModalEmpArea');
  if (filtroEmpArea) {
    filtroEmpArea.addEventListener('change', () => {
      renderModalEmployees();
    });
  }

  // Eventos del Modal Pasajeros No Contemplados en BD
  const btnOpenUnm = document.getElementById('btnOpenUnmappedModal');
  if (btnOpenUnm) {
    btnOpenUnm.addEventListener('click', () => {
      openUnmappedModal();
    });
  }

  const btnCloseUnm = document.getElementById('close-modal-no-contemplados');
  if (btnCloseUnm) {
    btnCloseUnm.addEventListener('click', () => {
      document.getElementById('modal-no-contemplados')?.classList.add('hidden');
    });
  }

  const modalUnm = document.getElementById('modal-no-contemplados');
  if (modalUnm) {
    modalUnm.addEventListener('click', (e) => {
      if (e.target === modalUnm) modalUnm.classList.add('hidden');
    });
  }

  const inputBuscarUnm = document.getElementById('inputBuscarNoContemplado');
  if (inputBuscarUnm) {
    inputBuscarUnm.addEventListener('input', () => {
      renderModalUnmapped();
    });
  }

  const btnCopiarDnis = document.getElementById('btnCopiarDnisNoMapeados');
  if (btnCopiarDnis) {
    btnCopiarDnis.addEventListener('click', () => {
      copiarDnisNoMapeados();
    });
  }

  // Pestañas de Navegación del Dashboard
  const btnTabOp = document.getElementById('btnTabOperacion');
  if (btnTabOp) {
    btnTabOp.addEventListener('click', () => switchDashboardTab('operacion'));
  }
  const btnTabCo = document.getElementById('btnTabCostos');
  if (btnTabCo) {
    btnTabCo.addEventListener('click', () => switchDashboardTab('costos'));
  }

  // Selector de Semana en Análisis de Costos
  const selSemCostos = document.getElementById('selectSemanaCostos');
  if (selSemCostos) {
    selSemCostos.addEventListener('change', (e) => {
      AppState.semanaCostosSeleccionada = e.target.value;
      renderAnalisisCostos();
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
    // Prioridad 1: Pestaña oficial de paraderos y rutas desde Google Drive
    const nameParaderos = getTabName(['RUTAS MOVILIDAD']) || 
                          getTabName(['RUTAS', 'MOVILIDAD']) || 
                          getTabName(['RUTAS_MOVILIDAD']) || 
                          getTabName(['PARADEROS_VALIDOS']) || 
                          getTabName(['PARADEROS VALIDOS']) || 
                          getTabName(['BD PARADEROS']) || 
                          getTabName(['PARADEROS']) || 
                          getTabName(['PARADERO']) || 
                          getTabName(['COORDENADAS']) || 
                          getTabName(['BD RUTAS']) || 
                          getTabName(['MIS RUTAS']) || 
                          (tabs.find(t => /^RUTAS?(\s+.*)?$/i.test(t.title.trim()))?.title);

    const [secos, ppa, frescos, registroDiario, paraderosSheet] = await Promise.all([
      nameSecos ? GoogleSheetsService.fetchSheetData(sheetId, nameSecos).catch(e => null) : Promise.resolve(null),
      namePpa ? GoogleSheetsService.fetchSheetData(sheetId, namePpa).catch(e => null) : Promise.resolve(null),
      nameFrescos ? GoogleSheetsService.fetchSheetData(sheetId, nameFrescos).catch(e => null) : Promise.resolve(null),
      nameRegistro ? GoogleSheetsService.fetchSheetData(sheetId, nameRegistro).catch(e => null) : Promise.resolve(null),
      nameParaderos ? GoogleSheetsService.fetchSheetData(sheetId, nameParaderos).catch(e => null) : Promise.resolve(null)
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
        const nombreRaw = getRowVal(row, ['APELLIDOS Y NOMBRES', 'NOMBRE Y APELLIDO', 'APELLIDOS Y NOMBRE', 'NOMBRE COMPLETO', 'COLABORADOR', 'EMPLEADO', 'TRABAJADOR', 'NOMBRE', 'NOMBRES', 'NAME']);
        const nombre = String(nombreRaw || '').trim() || 'Sin registrar';
        
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
          nombre: nombre,
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

    // Procesar datos de paraderos directamente desde el Google Sheet (100% dinámico)
    if (paraderosSheet && paraderosSheet.rows && paraderosSheet.rows.length > 0) {
      const parsedParaderos = [];
      paraderosSheet.rows.forEach((row, idx) => {
        let rutaVal = String(getRowVal(row, ['RUTA', 'LINEA', 'CODIGO RUTA', 'COD_RUTA', 'ID RUTA', 'ROUTE', 'RUTA PARADERO']) || '').trim();
        rutaVal = rutaVal.replace(/^RUTA\s+/i, '');
        const rawLat = getRowVal(row, ['LAT', 'LATITUD', 'LATITUDE', 'Y', 'COORD Y', 'COORDENADA Y', 'LAT-TRAB', 'LAT PARADERO']);
        const rawLng = getRowVal(row, ['LNG', 'LON', 'LONG', 'LONGITUD', 'LONGITUDE', 'X', 'COORD X', 'COORDENADA X', 'LON-TRAB', 'LON PARADERO']);
        const rawSec = getRowVal(row, ['SECUENCIA', 'ORDEN', 'PASO', 'NUMERO', 'ITEM', 'SEQ', 'ORD', 'NRO', 'N°']);
        const secuencia = parseInt(rawSec, 10) || (idx + 1);
        const nombre = String(getRowVal(row, ['NOMBRE', 'PARADERO', 'NOMBRE PARADERO', 'NOMBRE_PARADERO', 'PARADERO MÁS CERCANO', 'PARADERO MAS CERCANO', 'DESCRIPCION', 'PUNTO', 'ESTACION', 'STOP_NAME', 'REFERENCIA']) || '').trim() || `Paradero ${secuencia}`;

        if (rutaVal && rawLat && rawLng) {
          const lat = parseFloat(String(rawLat).replace(',', '.'));
          const lng = parseFloat(String(rawLng).replace(',', '.'));
          if (!isNaN(lat) && !isNaN(lng)) {
            parsedParaderos.push({ ruta: rutaVal, lat, lng, nombre, secuencia });
          }
        }
      });

      if (parsedParaderos.length > 0) {
        parsedParaderos.sort((a, b) => {
          if (a.ruta !== b.ruta) return a.ruta.localeCompare(b.ruta, undefined, { numeric: true });
          return a.secuencia - b.secuencia;
        });
        AppState.paraderosData = parsedParaderos;
        AppState.paraderosSource = 'DRIVE';
        AppState.paraderosTabName = nameParaderos;
        updateParaderosSourceBadge();
        AppState.mapDrawn = false;
        if (AppState.mapInstance) {
          dibujarRutasEnMapa();
        }
      } else {
        AppState.paraderosData = [];
        AppState.paraderosSource = 'DRIVE';
        AppState.paraderosTabName = '';
        updateParaderosSourceBadge();
      }
    } else {
      AppState.paraderosData = [];
      AppState.paraderosSource = 'DRIVE';
      AppState.paraderosTabName = '';
      updateParaderosSourceBadge();
    }

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

    // Identificar pasajeros en REGISTRO_DIARIO que no existen en las bases maestras
    const unmappedMap = new Map();
    AppState.rawRegistroDiario.forEach(row => {
      const rawDni = getRowVal(row, ['DNI', 'USERID', 'USER ID', 'DOCUMENTO', 'ID', 'CODIGO']);
      const cleanPassengerDni = cleanDni(rawDni);
      if (!cleanPassengerDni && !rawDni) return;

      const emp = AppState.employeeMap.get(cleanPassengerDni) || (rawDni ? AppState.employeeMap.get(String(rawDni).trim().toUpperCase()) : null);

      if (!emp) {
        const key = cleanPassengerDni || String(rawDni).trim().toUpperCase();
        const rutaVal = String(getRowVal(row, ['RUTA', 'RUTA ASIGNADA', 'LINEA']) || '').trim();
        const rawFecha = getRowVal(row, ['FECHA', 'FECHA DE VIAJE', 'DATE', 'DIA FECHA']);
        const fechaVal = normalizeDateStr(rawFecha);

        if (!unmappedMap.has(key)) {
          unmappedMap.set(key, {
            dni: key,
            rawDni: String(rawDni || key).trim(),
            viajesCount: 0,
            rutas: new Set(),
            fechas: new Set(),
            ultimaFecha: ''
          });
        }
        const item = unmappedMap.get(key);
        item.viajesCount += 1;
        if (rutaVal) item.rutas.add(rutaVal);
        if (fechaVal) {
          item.fechas.add(fechaVal);
          item.ultimaFecha = fechaVal;
        }
      }
    });

    AppState.unmappedPassengers = Array.from(unmappedMap.values()).sort((a, b) => b.viajesCount - a.viajesCount);

    // Actualizar contadores en botones de la barra de filtros y modales
    const bCountEmp = document.getElementById('badgeCountEmployees');
    if (bCountEmp) bCountEmp.innerText = combined.length;

    const bCountUnm = document.getElementById('badgeCountUnmapped');
    if (bCountUnm) bCountUnm.innerText = AppState.unmappedPassengers.length;

    const bInfoSinBD = document.getElementById('badgeInfoSinBD');
    if (bInfoSinBD) bInfoSinBD.innerText = `${AppState.unmappedPassengers.length} sin BD`;

    const mCountEmp = document.getElementById('modalEmpCountBadge');
    if (mCountEmp) mCountEmp.innerText = `${combined.length} colaboradores`;

    const mCountUnm = document.getElementById('modalUnmappedCountBadge');
    if (mCountUnm) mCountUnm.innerText = `${AppState.unmappedPassengers.length} sin BD`;

    document.getElementById('dashboardFiltersWrapper').classList.remove('hidden');
    document.getElementById('dashboardContent').classList.remove('hidden');

    applyFilters();
    renderAnalisisCostos();

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
  const isSegmented = filterAreaActive || filterTipoActive;

  // Opción del usuario: Incluir o excluir pasajeros sin BD maestra
  const chipSinBDEl = document.querySelector('#chipIncluirSinBD .chip.active');
  const incluirSinBD = chipSinBDEl ? chipSinBDEl.dataset.value : 'SI';

  // Si hay segmentación por Área o Tipo, o si el usuario eligió "NO" (Excluir sin BD), filtramos por BD maestra
  const isFiltered = isSegmented || (incluirSinBD === 'NO');

  // Actualizar UI del grupo de filtro sin BD
  const groupSinBDEl = document.getElementById('groupFiltroSinBD');
  const labelSinBDHint = document.getElementById('labelSinBDHint');
  if (groupSinBDEl) {
    if (isSegmented) {
      groupSinBDEl.style.opacity = '0.55';
      groupSinBDEl.title = 'Al filtrar por Área o Tipo, los pasajeros sin BD ya quedan excluidos automáticamente al no tener área ni tipo asignado.';
      if (labelSinBDHint) {
        labelSinBDHint.innerText = '(Excluidos por filtro)';
        labelSinBDHint.style.color = '#94a3b8';
      }
    } else {
      groupSinBDEl.style.opacity = '1';
      groupSinBDEl.title = '';
      if (labelSinBDHint) {
        if (incluirSinBD === 'SI') {
          labelSinBDHint.innerText = '(Incluidos en "TODOS")';
          labelSinBDHint.style.color = '#4ade80';
        } else {
          labelSinBDHint.innerText = '(Excluidos - Solo BD)';
          labelSinBDHint.style.color = '#f59e0b';
        }
      }
    }
  }

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
  let aggrTripIdx = 0;
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
    const rawRowDni = getRowVal(row, ['DNI', 'USERID', 'USER ID', 'DOCUMENTO', 'ID', 'CODIGO']);
    const isPassengerRow = Boolean(rawRowDni && String(rawRowDni).trim().length >= 4);
    const tripKey = isPassengerRow ? `${fechaSoloDia}|${rutaVal}|${tipoBusVal || 'BUS'}` : `${fechaSoloDia}|${rutaVal}|${tipoBusVal || 'BUS'}|${aggrTripIdx++}`;

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
    const rawPasajCol = getRowVal(row, ['PASAJEROS', 'TOTAL PASAJEROS', 'PASAJ', 'CANTIDAD PASAJEROS', 'CANT_PASAJEROS']);
    const numPasajCol = parseFloat(String(rawPasajCol || '0').replace(/[^0-9.-]+/g, "")) || 0;

    if (isPassengerRow) {
      trip.totalPasajeros += 1;
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
    } else {
      const cant = numPasajCol > 0 ? numPasajCol : 1;
      trip.totalPasajeros += cant;
      trip.pasajerosFiltrados += cant;
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
  renderAnalisisCostos();
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
  // Variables para tablas de clasificación con rangos de distancia explícitos
  const countCd = {
    '🟢 Muy Cerca (≤ 5 km)': 0,
    '🟡 Cerca (5 - 10 km)': 0,
    '🟠 Moderada (10 - 20 km)': 0,
    '🔴 Lejos (> 20 km)': 0
  };

  const countParadero = {
    '🟢 Muy Cerca (≤ 1 km)': 0,
    '🟡 Cerca (1 - 3 km)': 0,
    '🟠 Moderada (3 - 5 km)': 0,
    '🔴 Lejos (> 5 km)': 0
  };

  const distritosStats = {};
  const rutasCount = {};
  const paraderosCount = {};

  data.forEach(emp => {
    // Totales
    if (emp.distCd > 0) { sumDistCd += emp.distCd; countValidCd++; }
    if (emp.distParadero > 0) { sumDistParadero += emp.distParadero; countValidParadero++; }
    if (emp.distrito) distritosSet.add(emp.distrito);
    if (emp.ruta) rutasSet.add(emp.ruta);

    // Clasificaciones CD (matching substring o valor numérico)
    if (emp.clasifCd.includes('Muy Cerca') || (emp.distCd > 0 && emp.distCd <= 5)) countCd['🟢 Muy Cerca (≤ 5 km)']++;
    else if (emp.clasifCd.includes('Cerca') || (emp.distCd > 5 && emp.distCd <= 10)) countCd['🟡 Cerca (5 - 10 km)']++;
    else if (emp.clasifCd.includes('Moderada') || (emp.distCd > 10 && emp.distCd <= 20)) countCd['🟠 Moderada (10 - 20 km)']++;
    else if (emp.clasifCd.includes('Lejos') || emp.distCd > 20) countCd['🔴 Lejos (> 20 km)']++;

    // Clasificaciones Paradero (matching substring o valor numérico)
    if (emp.clasifParadero.includes('Muy Cerca') || (emp.distParadero > 0 && emp.distParadero <= 1)) countParadero['🟢 Muy Cerca (≤ 1 km)']++;
    else if (emp.clasifParadero.includes('Cerca') || (emp.distParadero > 1 && emp.distParadero <= 3)) countParadero['🟡 Cerca (1 - 3 km)']++;
    else if (emp.clasifParadero.includes('Moderada') || (emp.distParadero > 3 && emp.distParadero <= 5)) countParadero['🟠 Moderada (3 - 5 km)']++;
    else if (emp.clasifParadero.includes('Lejos') || emp.distParadero > 5) countParadero['🔴 Lejos (> 5 km)']++;

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
    const subTotalEmps = document.querySelector('#kpiTotalEmps + .kpi-subtitle');
    if (subTotalEmps) {
      const chipSinBDEl = document.querySelector('#chipIncluirSinBD .chip.active');
      const incluirSinBDActive = chipSinBDEl ? chipSinBDEl.dataset.value !== 'NO' : true;
      const areasAct = Array.from(document.querySelectorAll('#chipArea .chip.active')).map(b => b.dataset.value);
      const tiposAct = Array.from(document.querySelectorAll('#chipTipo .chip.active')).map(b => b.dataset.value);
      const isSeg = (!areasAct.includes('TODOS') && areasAct.length > 0) || (!tiposAct.includes('TODOS') && tiposAct.length > 0);

      if (!isSeg && incluirSinBDActive && (AppState.unmappedPassengers || []).length > 0) {
        subTotalEmps.innerText = `En BD (+${AppState.unmappedPassengers.length} sin registrar)`;
      } else {
        subTotalEmps.innerText = 'Personal auditado en BD';
      }
    }
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
  const colorsCd = {
    '🟢 Muy Cerca (≤ 5 km)': '#22c55e', 
    '🟡 Cerca (5 - 10 km)': '#eab308',     
    '🟠 Moderada (10 - 20 km)': '#f97316',  
    '🔴 Lejos (> 20 km)': '#ef4444'      
  };

  const colorsParadero = {
    '🟢 Muy Cerca (≤ 1 km)': '#22c55e', 
    '🟡 Cerca (1 - 3 km)': '#eab308',     
    '🟠 Moderada (3 - 5 km)': '#f97316',  
    '🔴 Lejos (> 5 km)': '#ef4444'      
  };

  renderPieChart('chartDistCD', countCd, colorsCd);
  renderPieChart('chartDistParadero', countParadero, colorsParadero);

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
    const chipSinBDEl = document.querySelector('#chipIncluirSinBD .chip.active');
    const incluirSinBDActive = chipSinBDEl ? chipSinBDEl.dataset.value !== 'NO' : true;
    const areasAct = Array.from(document.querySelectorAll('#chipArea .chip.active')).map(b => b.dataset.value);
    const tiposAct = Array.from(document.querySelectorAll('#chipTipo .chip.active')).map(b => b.dataset.value);
    const isSeg = (!areasAct.includes('TODOS') && areasAct.length > 0) || (!tiposAct.includes('TODOS') && tiposAct.length > 0);

    const subCosto = document.querySelector('#kpiCostoDia + .kpi-subtitle');
    if (subCosto) {
      if (!isSeg && !incluirSinBDActive) {
        subCosto.innerText = isVerTodas ? 'Total invertido (Solo BD)' : 'Invertido hoy (Solo BD)';
      } else {
        subCosto.innerText = isVerTodas ? 'Total invertido acumulado' : 'Total invertido hoy';
      }
    }

    const subOcup = document.querySelector('#kpiOcupacionDia + .kpi-subtitle');
    if (subOcup) {
      if (!isSeg && !incluirSinBDActive) {
        subOcup.innerText = isVerTodas ? 'Eficiencia (Solo BD)' : 'Eficiencia hoy (Solo BD)';
      } else {
        subOcup.innerText = isVerTodas ? 'Eficiencia acumulada' : 'Eficiencia operativa hoy';
      }
    }

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
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const val = context.parsed || 0;
              const dataset = context.dataset.data || [];
              const total = dataset.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? ((val / total) * 100).toFixed(1) + '%' : '0%';
              return ` ${context.label}: ${val} colaboradores (${pct})`;
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

// ==========================================
// Módulo de Mapeo de Rutas, Paraderos y GPS
// ==========================================

function updateParaderosSourceBadge() {
  const badge = document.getElementById('mapaDataSourceBadge') || document.getElementById('mapaOrigenBadge');
  const badgePill = document.getElementById('badgeRutasMap');
  const count = (AppState.paraderosData || []).length;
  
  if (count > 0) {
    if (badge) {
      badge.className = 'file-badge badge-owned';
      badge.style.background = 'rgba(34, 197, 94, 0.15)';
      badge.style.color = '#4ade80';
      badge.style.borderColor = 'rgba(34, 197, 94, 0.3)';
      badge.innerHTML = `<i class="fa-solid fa-cloud-check"></i> Google Drive: ${AppState.paraderosTabName || 'RUTAS MOVILIDAD'} (${count} pts)`;
    }
    if (badgePill) {
      badgePill.style.background = '#16a34a';
      badgePill.innerText = 'Drive';
    }
  } else {
    if (badge) {
      badge.className = 'file-badge';
      badge.style.background = 'rgba(148, 163, 184, 0.15)';
      badge.style.color = '#94a3b8';
      badge.style.borderColor = 'rgba(148, 163, 184, 0.3)';
      badge.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i> Google Drive: Pendiente de sincronización`;
    }
    if (badgePill) {
      badgePill.style.background = '#64748b';
      badgePill.innerText = 'Drive';
    }
  }
}

function openMapModal() {
  const modal = document.getElementById('mapa-modal');
  if (!modal) return;
  modal.classList.remove('hidden');

  updateParaderosSourceBadge();
  initRoutesMap();

  setTimeout(() => {
    if (AppState.mapInstance) {
      AppState.mapInstance.invalidateSize();
      if (AppState.mapBounds && AppState.mapBounds.isValid()) {
        AppState.mapInstance.fitBounds(AppState.mapBounds, { padding: [40, 40] });
      }
    }
  }, 150);

  if (!AppState.mapDrawn) {
    dibujarRutasEnMapa();
  }
}

function closeMapModal() {
  const modal = document.getElementById('mapa-modal');
  if (modal) modal.classList.add('hidden');
}

function initRoutesMap() {
  if (AppState.mapInstance) return;
  const mapContainer = document.getElementById('mapaRutas');
  if (!mapContainer) return;

  const map = L.map('mapaRutas', {
    center: [-12.046374, -77.042793],
    zoom: 11,
    wheelDebounceTime: 60,
    wheelPxPerZoomLevel: 100
  });

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© OpenStreetMap contributors',
    maxZoom: 19,
    keepBuffer: 6,
    updateWhenZooming: false,
    updateWhenIdle: true
  }).addTo(map);

  map.on('click', function (e) {
    if (e.originalEvent.target.id === 'mapaRutas' || e.originalEvent.target.classList.contains('leaflet-container')) {
      // Si el usuario hace clic en el mapa, quitamos el filtro de ruta sin cambiar el zoom ni la posición de la cámara
      const sel = document.getElementById('selectFiltroRutaMapa');
      if (sel && sel.value !== 'TODAS') {
        sel.value = 'TODAS';
        resetearVistaMapa(false); // recenter = false: ¡NUNCA aleja ni recentra!
      }
    }
  });

  AppState.mapInstance = map;
}

function dibujarRutasEnMapa() {
  const contenedorLista = document.getElementById('listaRutasCards');
  if (!AppState.paraderosData || AppState.paraderosData.length === 0) {
    if (contenedorLista) {
      contenedorLista.innerHTML = `
        <div style="text-align: center; color: #94a3b8; padding: 30px 15px;">
          <i class="fa-solid fa-triangle-exclamation" style="font-size: 1.8rem; color: #f59e0b; margin-bottom: 10px;"></i>
          <p>No se encontraron coordenadas de paraderos cargadas desde Google Drive.</p>
          <small>Asegúrate de tener la pestaña <b>RUTAS MOVILIDAD</b> o <b>PARADEROS_VALIDOS</b> en tu Google Sheet.</small>
        </div>
      `;
    }
    return;
  }

  // Limpiar capas previas si ya existían
  if (AppState.mapLines) {
    Object.values(AppState.mapLines).forEach(l => {
      if (l && l.remove) l.remove();
    });
  }
  if (AppState.mapMarkers) {
    Object.values(AppState.mapMarkers).forEach(arr => {
      arr.forEach(m => { if (m && m.remove) m.remove(); });
    });
  }

  AppState.mapLines = {};
  AppState.mapMarkers = {};
  AppState.mapCards = {};

  const rutasAgrupadas = {};
  const allLatLngs = [];

  AppState.paraderosData.forEach(punto => {
    let rId = String(punto.ruta || '').trim().toUpperCase();
    rId = rId.replace(/^RUTA\s+/i, '');
    if (!rId) return;

    if (!rutasAgrupadas[rId]) {
      rutasAgrupadas[rId] = [];
      AppState.mapMarkers[rId] = [];
    }

    const lat = parseFloat(punto.lat);
    const lng = parseFloat(punto.lng);
    if (isNaN(lat) || isNaN(lng)) return;

    rutasAgrupadas[rId].push({
      lat,
      lng,
      nombre: punto.nombre || `Paradero`,
      secuencia: parseInt(punto.secuencia, 10) || (rutasAgrupadas[rId].length + 1)
    });
    allLatLngs.push([lat, lng]);
  });

  // Ordenar paraderos por secuencia dentro de cada ruta
  Object.keys(rutasAgrupadas).forEach(rId => {
    rutasAgrupadas[rId].sort((a, b) => a.secuencia - b.secuencia);
  });

  // Poblar dropdown de rutas
  const selectRuta = document.getElementById('selectFiltroRutaMapa');
  const sortedRutas = Object.keys(rutasAgrupadas).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  if (selectRuta) {
    selectRuta.innerHTML = `<option value="TODAS">Todas las Rutas (${sortedRutas.length})</option>` +
      sortedRutas.map(r => `<option value="${r}">Ruta ${r}</option>`).join('');
    selectRuta.value = 'TODAS';
  }

  // Ajustar límites iniciales
  if (allLatLngs.length > 0 && AppState.mapInstance) {
    AppState.mapBounds = L.latLngBounds(allLatLngs);
    AppState.mapInstance.fitBounds(AppState.mapBounds, { padding: [40, 40] });
  }

  // Dibujar Pines con Tooltips eficientes (solo al hover o al seleccionar ruta)
  sortedRutas.forEach(rId => {
    const puntos = rutasAgrupadas[rId];
    const color = getRutaColor(rId);

    puntos.forEach(p => {
      const marcador = L.marker([p.lat, p.lng], { icon: crearIconoDePinColoreado(color) })
        .addTo(AppState.mapInstance)
        .bindTooltip(`<b>Ruta ${rId}</b>: ${p.nombre}`, {
          permanent: false, // Solo visible en hover o al filtrar ruta -> 100% fluido y sin sobrecargar la pantalla
          direction: 'top',
          className: 'etiqueta-paradero',
          offset: [0, -18]
        })
        .on('click', (e) => {
          L.DomEvent.stopPropagation(e);
          resaltarRutaEnMapa(rId);
        });

      AppState.mapMarkers[rId].push(marcador);
    });
  });

  if (contenedorLista) contenedorLista.innerHTML = '';

  // Solicitar trazado a Mapbox o reutilizar Caché
  sortedRutas.forEach(nombreRuta => {
    const puntos = rutasAgrupadas[nombreRuta];
    if (puntos.length < 2) return;

    const colorAsignado = getRutaColor(nombreRuta);
    const coordsGoogle = puntos.map(p => `${p.lat},${p.lng}`);

    let botonesGPS = '';
    if (coordsGoogle.length <= 10) {
      botonesGPS = `<a href="${generarLinkGoogle(coordsGoogle)}" target="_blank" rel="noopener" class="btn-gps"><i class="fa-solid fa-location-arrow"></i> Abrir en Google Maps</a>`;
    } else {
      const parte1 = coordsGoogle.slice(0, 10);
      const parte2 = coordsGoogle.slice(9);
      botonesGPS = `
        <div style="display:flex; gap:6px;">
          <a href="${generarLinkGoogle(parte1)}" target="_blank" rel="noopener" class="btn-gps" style="flex:1;"><i class="fa-solid fa-location-arrow"></i> GPS Parte 1</a>
          <a href="${generarLinkGoogle(parte2)}" target="_blank" rel="noopener" class="btn-gps" style="flex:1; background-color:#16a34a;"><i class="fa-solid fa-location-arrow"></i> GPS Parte 2</a>
        </div>
      `;
    }

    const tarjetaRuta = document.createElement('div');
    tarjetaRuta.className = 'tarjeta-logistica';
    tarjetaRuta.style.borderLeft = `5px solid ${colorAsignado}`;
    tarjetaRuta.dataset.ruta = nombreRuta;

    const listaParaderosHtml = `
      <div class="tarjeta-paraderos-preview" style="margin: 8px 0 10px 0; max-height: 110px; overflow-y: auto; background: rgba(15, 23, 42, 0.6); border-radius: 8px; padding: 6px 10px; border: 1px solid rgba(51, 65, 85, 0.4);">
        ${puntos.map(p => `
          <div style="font-size: 0.76rem; color: #cbd5e1; padding: 2px 0; display: flex; align-items: baseline; gap: 6px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
            <span style="color: ${colorAsignado}; font-weight: 700; font-size: 0.72rem; min-width: 16px;">${p.secuencia}.</span>
            <span title="${p.nombre}" style="overflow: hidden; text-overflow: ellipsis;">${p.nombre}</span>
          </div>
        `).join('')}
      </div>
    `;

    tarjetaRuta.innerHTML = `
      <div class="tarjeta-ruta-title" style="color: ${colorAsignado};">
        <span><i class="fa-solid fa-bus"></i> RUTA ${nombreRuta}</span>
        <span style="font-size: 0.75rem; color: #94a3b8; font-weight: normal;">${puntos.length} paraderos</span>
      </div>
      <div class="tarjeta-stats" id="stats-ruta-${nombreRuta}">
        <span><i class="fa-solid fa-road"></i> Calculando...</span>
        <span><i class="fa-regular fa-clock"></i> ...</span>
      </div>
      ${listaParaderosHtml}
      ${botonesGPS}
    `;

    tarjetaRuta.addEventListener('click', (e) => {
      if (e.target.tagName !== 'A' && !e.target.closest('a')) {
        resaltarRutaEnMapa(nombreRuta);
      }
    });

    if (contenedorLista) contenedorLista.appendChild(tarjetaRuta);
    AppState.mapCards[nombreRuta] = tarjetaRuta;

    // Helpers para aplicar trazado en mapa
    const aplicarTrazadoGeoJSON = (rutaData) => {
      const distKM = (rutaData.distance / 1000).toFixed(1);
      const durMin = Math.round(rutaData.duration / 60);

      const statsEl = document.getElementById(`stats-ruta-${nombreRuta}`);
      if (statsEl) {
        statsEl.innerHTML = `<span><i class="fa-solid fa-road"></i> ${distKM} km</span><span><i class="fa-regular fa-clock"></i> ~${durMin} min</span>`;
      }

      const layer = L.geoJSON(rutaData.geometry, {
        style: { color: colorAsignado, weight: 4, opacity: 0.8 }
      }).addTo(AppState.mapInstance);

      layer.on('click', (e) => {
        L.DomEvent.stopPropagation(e);
        resaltarRutaEnMapa(nombreRuta);
      });

      AppState.mapLines[nombreRuta] = layer;
    };

    const aplicarTrazadoDirecto = () => {
      const latLngs = puntos.map(p => [p.lat, p.lng]);
      const layer = L.polyline(latLngs, {
        color: colorAsignado,
        weight: 4,
        opacity: 0.8,
        dashArray: '6, 6'
      }).addTo(AppState.mapInstance);

      layer.on('click', (e) => {
        L.DomEvent.stopPropagation(e);
        resaltarRutaEnMapa(nombreRuta);
      });

      AppState.mapLines[nombreRuta] = layer;

      const statsEl = document.getElementById(`stats-ruta-${nombreRuta}`);
      if (statsEl) {
        statsEl.innerHTML = `<span><i class="fa-solid fa-road"></i> Trazado directo</span><span><i class="fa-regular fa-compass"></i> ${puntos.length} paraderos</span>`;
      }
    };

    const stringCoordenadas = puntos.map(p => `${p.lng},${p.lat}`).join(';');
    const cacheKey = nombreRuta + '_' + stringCoordenadas;

    // Si ya lo tenemos en caché, dibujar al instante (0ms)
    if (AppState.routeDirectionsCache && AppState.routeDirectionsCache[cacheKey]) {
      aplicarTrazadoGeoJSON(AppState.routeDirectionsCache[cacheKey]);
    } else {
      const urlMapbox = `https://api.mapbox.com/directions/v5/mapbox/driving/${stringCoordenadas}?geometries=geojson&access_token=${AppState.mapboxToken}`;
      fetch(urlMapbox)
        .then(r => r.json())
        .then(data => {
          if (data.code === 'Ok' && data.routes && data.routes.length > 0) {
            const rutaData = data.routes[0];
            if (AppState.routeDirectionsCache) {
              AppState.routeDirectionsCache[cacheKey] = rutaData;
            }
            aplicarTrazadoGeoJSON(rutaData);
          } else {
            aplicarTrazadoDirecto();
          }
        })
        .catch(() => {
          aplicarTrazadoDirecto();
        });
    }
  });

  AppState.mapDrawn = true;
}

function resaltarRutaEnMapa(idRutaSeleccionada) {
  if (!idRutaSeleccionada || idRutaSeleccionada === 'TODAS') {
    resetearVistaMapa(false);
    return;
  }

  // 1. Sincronizar select
  const sel = document.getElementById('selectFiltroRutaMapa');
  if (sel && sel.value !== idRutaSeleccionada) {
    sel.value = idRutaSeleccionada;
  }

  // 2. Líneas del mapa
  for (const [idRuta, linea] of Object.entries(AppState.mapLines)) {
    const color = getRutaColor(idRuta);
    if (idRuta === idRutaSeleccionada) {
      linea.setStyle({ weight: 7, color: color, opacity: 1 });
      if (linea.bringToFront) linea.bringToFront();
    } else {
      linea.setStyle({ weight: 2, color: color, opacity: 0.15 });
    }
  }

  // 3. Marcadores y Tooltips: abrir etiquetas SOLO para la ruta seleccionada
  for (const [idRuta, marcadores] of Object.entries(AppState.mapMarkers)) {
    if (idRuta === idRutaSeleccionada) {
      marcadores.forEach(m => {
        m.setOpacity(1);
        m.setZIndexOffset(1000);
        m.openTooltip(); // Muestra el nombre solo para los paraderos de esta ruta
        const tooltipEl = m.getTooltip()?.getElement();
        if (tooltipEl) tooltipEl.style.opacity = '1';
      });
    } else {
      marcadores.forEach(m => {
        m.setOpacity(0.18);
        m.setZIndexOffset(0);
        m.closeTooltip(); // Oculta el nombre de las otras rutas
        const tooltipEl = m.getTooltip()?.getElement();
        if (tooltipEl) tooltipEl.style.opacity = '0.15';
      });
    }
  }

  // 4. Tarjetas del panel
  document.querySelectorAll('.tarjeta-logistica').forEach(t => t.classList.remove('tarjeta-resaltada'));
  const card = AppState.mapCards[idRutaSeleccionada];
  if (card) {
    card.classList.add('tarjeta-resaltada');
    card.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // 5. Ajustar vista a la ruta
  if (AppState.mapLines[idRutaSeleccionada] && AppState.mapInstance) {
    const bounds = AppState.mapLines[idRutaSeleccionada].getBounds();
    if (bounds && bounds.isValid()) {
      AppState.mapInstance.fitBounds(bounds, { padding: [50, 50], maxZoom: 14 });
    }
  }
}

function resetearVistaMapa(recenter = false) {
  // Restaurar líneas
  for (const [id, linea] of Object.entries(AppState.mapLines)) {
    linea.setStyle({ weight: 4, color: getRutaColor(id), opacity: 0.8 });
  }

  // Restaurar paneles
  document.querySelectorAll('.tarjeta-resaltada').forEach(t => t.classList.remove('tarjeta-resaltada'));

  // Restaurar marcadores y cerrar tooltips
  for (const [id, marcadores] of Object.entries(AppState.mapMarkers)) {
    marcadores.forEach(m => {
      m.setOpacity(1);
      m.setZIndexOffset(0);
      m.closeTooltip(); // Cierra tooltip para que el mapa quede despejado y veloz
      const tooltipEl = m.getTooltip()?.getElement();
      if (tooltipEl) tooltipEl.style.opacity = '1';
    });
  }

  // Sincronizar select
  const sel = document.getElementById('selectFiltroRutaMapa');
  if (sel) sel.value = 'TODAS';

  // Solo re-centrar el mapa si se solicitó explícitamente desde el botón [Centrar]
  if (recenter && AppState.mapBounds && AppState.mapBounds.isValid() && AppState.mapInstance) {
    AppState.mapInstance.fitBounds(AppState.mapBounds, { padding: [40, 40] });
  }
}

// ==========================================
// Módulo de Directorio de Colaboradores
// ==========================================

function openEmployeesModal() {
  const modal = document.getElementById('modal-empleados');
  if (!modal) return;
  modal.classList.remove('hidden');
  renderModalEmployees();
}

function renderModalEmployees() {
  const tbody = document.getElementById('tbodyModalEmpleados');
  if (!tbody) return;

  const search = String(document.getElementById('inputBuscarEmpleado')?.value || '').trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const filtroTipo = document.getElementById('filtroModalEmpTipo')?.value || 'TODOS';
  const filtroArea = document.getElementById('filtroModalEmpArea')?.value || 'TODOS';

  const emps = (AppState.rawEmployees || []).filter(e => {
    if (filtroTipo !== 'TODOS' && e.tipo !== filtroTipo) return false;
    if (filtroArea !== 'TODOS' && e.area !== filtroArea) return false;
    if (!search) return true;

    const dniNorm = (e.dni || '').toLowerCase();
    const rawDniNorm = (e.rawDni || '').toLowerCase();
    const nomNorm = (e.nombre || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const distNorm = (e.distrito || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const rutaNorm = (e.ruta || '').toLowerCase();

    return dniNorm.includes(search) || rawDniNorm.includes(search) || nomNorm.includes(search) || distNorm.includes(search) || rutaNorm.includes(search);
  });

  const footer = document.getElementById('labelTotalMostradosEmp');
  if (footer) {
    footer.innerText = `Mostrando ${emps.length} de ${(AppState.rawEmployees || []).length} colaboradores`;
  }

  if (emps.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: #94a3b8; padding: 25px;">No se encontraron colaboradores con los filtros seleccionados.</td></tr>`;
    return;
  }

  // Renderizar los primeros 300 para asegurar fluidez instantánea
  const limit = 300;
  const slice = emps.slice(0, limit);

  tbody.innerHTML = slice.map(e => {
    const areaBadge = e.area === 'SECOS' 
      ? '<span class="file-badge badge-owned">SECOS</span>' 
      : (e.area === 'PPA' ? '<span class="file-badge badge-shared">PPA</span>' : '<span class="file-badge" style="background: rgba(6, 182, 212, 0.15); color: #22d3ee; border: 1px solid rgba(6, 182, 212, 0.3);">FRESCOS</span>');
    
    const tipoBadge = e.tipo === 'STAFF'
      ? '<span class="badge badge-info" style="font-size: 0.72rem;">STAFF</span>'
      : (e.tipo === 'OPERARIO' ? '<span class="badge badge-success" style="font-size: 0.72rem;">OPERARIO</span>' : `<span class="badge" style="background: #334155; color: #94a3b8; font-size: 0.72rem;">${e.tipo}</span>`);

    let rutaDisplay = '-';
    if (e.ruta) {
      const rClean = String(e.ruta).trim();
      rutaDisplay = rClean.toUpperCase().startsWith('RUTA') ? rClean : `Ruta ${rClean}`;
    }

    return `<tr>
      <td style="font-weight: 600; font-family: monospace; color: #f8fafc; white-space: nowrap;">${e.dni || e.rawDni}</td>
      <td style="color: #cbd5e1; font-weight: 500; white-space: nowrap;">${e.nombre || 'Sin registrar'}</td>
      <td style="text-align: center; white-space: nowrap;">${areaBadge}</td>
      <td style="text-align: center; white-space: nowrap;">${tipoBadge}</td>
      <td style="color: #94a3b8; white-space: nowrap;">${e.distrito || '-'}</td>
      <td style="color: #60a5fa; font-weight: 600; white-space: nowrap;">${rutaDisplay}</td>
      <td style="color: #cbd5e1; font-size: 0.85rem; white-space: nowrap;">${e.paradero || '-'}</td>
    </tr>`;
  }).join('');

  if (emps.length > limit) {
    tbody.innerHTML += `<tr><td colspan="7" style="text-align: center; color: #f59e0b; padding: 10px; font-size: 0.8rem;">Mostrando los primeros ${limit} de ${emps.length} resultados. Usa el buscador para filtrar más específicamente.</td></tr>`;
  }
}

// ==========================================
// Módulo de Pasajeros No Contemplados en BD
// ==========================================

function openUnmappedModal() {
  const modal = document.getElementById('modal-no-contemplados');
  if (!modal) return;
  modal.classList.remove('hidden');
  renderModalUnmapped();
}

function renderModalUnmapped() {
  const tbody = document.getElementById('tbodyModalNoContemplados');
  if (!tbody) return;

  const search = String(document.getElementById('inputBuscarNoContemplado')?.value || '').trim().toLowerCase();

  const list = (AppState.unmappedPassengers || []).filter(u => {
    if (!search) return true;
    const dniNorm = (u.dni || '').toLowerCase();
    const rawDniNorm = (u.rawDni || '').toLowerCase();
    const rutasNorm = Array.from(u.rutas).join(' ').toLowerCase();
    return dniNorm.includes(search) || rawDniNorm.includes(search) || rutasNorm.includes(search);
  });

  const footer = document.getElementById('labelTotalMostradosUnmapped');
  if (footer) {
    footer.innerText = `Total: ${list.length} pasajeros en Registro Diario sin registro en BD maestras`;
  }

  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #4ade80; padding: 25px;"><i class="fa-solid fa-circle-check"></i> Excelente. Todos los pasajeros registrados coinciden con las bases maestras.</td></tr>`;
    return;
  }

  tbody.innerHTML = list.map(u => {
    const rutasStr = Array.from(u.rutas).map(r => {
      const rClean = String(r).trim();
      const rDisp = rClean.toUpperCase().startsWith('RUTA') ? rClean : `Ruta ${rClean}`;
      return `<span class="badge" style="background: #1e293b; border: 1px solid #334155; color: #93c5fd; margin-right: 4px; white-space: nowrap;">${rDisp}</span>`;
    }).join('') || '-';

    return `<tr>
      <td style="font-weight: 700; font-family: monospace; color: #f87171; white-space: nowrap;">${u.dni || u.rawDni}</td>
      <td style="font-weight: 600; color: #f8fafc; text-align: center; white-space: nowrap;">${u.viajesCount}</td>
      <td style="white-space: nowrap;">${rutasStr}</td>
      <td style="color: #94a3b8; font-size: 0.82rem; white-space: nowrap;">${u.ultimaFecha || '-'}</td>
      <td style="white-space: nowrap;"><span class="file-badge" style="background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); white-space: nowrap; display: inline-block;">🔴 No figura en BD SECOS/PPA/FRESCOS</span></td>
    </tr>`;
  }).join('');
}

function copiarDnisNoMapeados() {
  const list = AppState.unmappedPassengers || [];
  if (list.length === 0) {
    if (window.ClipboardUtil) ClipboardUtil.showToast('No hay DNIs sin mapear para copiar', 'info');
    else alert('No hay DNIs sin mapear para copiar');
    return;
  }

  const text = list.map(u => u.dni || u.rawDni).join('\n');
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => {
      if (window.ClipboardUtil) ClipboardUtil.showToast(`Se copiaron ${list.length} DNIs al portapapeles`, 'success');
      else alert(`Se copiaron ${list.length} DNIs al portapapeles`);
    });
  } else {
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    if (window.ClipboardUtil) ClipboardUtil.showToast(`Se copiaron ${list.length} DNIs al portapapeles`, 'success');
    else alert(`Se copiaron ${list.length} DNIs al portapapeles`);
  }
}

// =========================================================
// MÓDULO: Pestañas de Navegación del Dashboard
// =========================================================

function switchDashboardTab(tabName) {
  AppState.activeTab = tabName;
  const btnOp = document.getElementById('btnTabOperacion');
  const btnCo = document.getElementById('btnTabCostos');
  const viewOp = document.getElementById('viewOperacionDemografia');
  const viewCo = document.getElementById('viewAnalisisCostos');

  if (tabName === 'costos') {
    if (btnOp) btnOp.classList.remove('active');
    if (btnCo) btnCo.classList.add('active');
    if (viewOp) viewOp.classList.add('hidden');
    if (viewCo) viewCo.classList.remove('hidden');
    renderAnalisisCostos();
    setTimeout(() => {
      if (AppState.charts && AppState.charts['chartOcupacionRutasCostos']) {
        AppState.charts['chartOcupacionRutasCostos'].resize();
      }
      if (AppState.charts && AppState.charts['chartCostoPasajerosRutas']) {
        AppState.charts['chartCostoPasajerosRutas'].resize();
      }
    }, 60);
  } else {
    if (btnCo) btnCo.classList.remove('active');
    if (btnOp) btnOp.classList.add('active');
    if (viewCo) viewCo.classList.add('hidden');
    if (viewOp) viewOp.classList.remove('hidden');
  }
}

// =========================================================
// MÓDULO: Hoja de Análisis de Costos y Eficiencia
// =========================================================

function renderAnalisisCostos() {
  const rawRows = AppState.rawRegistroDiario || [];
  if (rawRows.length === 0) return;

  // 1. Extraer semanas únicas disponibles
  const semSet = new Set();
  rawRows.forEach(r => {
    const s = parseInt(getRowVal(r, ['SEMANA', 'SEM']), 10);
    if (!isNaN(s) && s > 0) semSet.add(s);
  });
  const semanas = Array.from(semSet).sort((a, b) => a - b);

  // 2. Poblar selector de semana si no tiene opciones suficientes
  const selSem = document.getElementById('selectSemanaCostos');
  if (selSem) {
    if (selSem.options.length <= 1 && semanas.length > 0) {
      selSem.innerHTML = semanas.map(s => `<option value="${s}">Semana ${s}</option>`).join('');
      if (semanas.length > 1) {
        selSem.innerHTML += `<option value="TODAS">Todas las Semanas</option>`;
      }
    }
    if (!AppState.semanaCostosSeleccionada) {
      AppState.semanaCostosSeleccionada = semanas.length > 0 ? String(semanas[semanas.length - 1]) : '9';
    }
    selSem.value = AppState.semanaCostosSeleccionada;
  }

  const semSel = AppState.semanaCostosSeleccionada || (semanas.length > 0 ? String(semanas[semanas.length - 1]) : '9');
  const isVerTodas = (semSel === 'TODAS');

  // Actualizar badge de la tabla
  const badgeSem = document.getElementById('badgeSemanaTabla');
  if (badgeSem) {
    badgeSem.innerText = isVerTodas ? 'Todas las Semanas' : `Semana ${semSel}`;
  }

  // 3. Filtrar registros por la semana analizada
  const rowsSemana = isVerTodas 
    ? rawRows 
    : rawRows.filter(r => {
        const s = parseInt(getRowVal(r, ['SEMANA', 'SEM']), 10);
        return String(s) === String(semSel);
      });

  // 4. Agregar métricas por ruta de la semana
  const statsRuta = {};
  rowsSemana.forEach(r => {
    const rutaRaw = String(getRowVal(r, ['RUTA', 'RUTA ASIGNADA', 'LINEA']) || '').trim();
    if (!rutaRaw) return;
    const ruta = rutaRaw.toUpperCase().startsWith('RUTA') ? rutaRaw.toUpperCase() : `RUTA ${rutaRaw.toUpperCase()}`;

    if (!statsRuta[ruta]) {
      statsRuta[ruta] = { viajes: 0, pasajeros: 0, capacidad: 0, costo: 0 };
    }
    const st = statsRuta[ruta];
    st.viajes += 1;

    const rawCap = getRowVal(r, ['CAPACIDAD', 'CAPACIDAD DE BUS', 'CAPACIDAD BUS', 'CAPACIDAD_BUS']);
    const capNum = parseFloat(String(rawCap || '0').replace(/[^0-9.-]+/g, "")) || 0;
    st.capacidad += capNum;

    const rawPasaj = getRowVal(r, ['PASAJEROS', 'TOTAL PASAJEROS', 'PASAJ', 'CANTIDAD PASAJEROS']);
    const numPasaj = parseFloat(String(rawPasaj || '0').replace(/[^0-9.-]+/g, "")) || 0;
    st.pasajeros += (numPasaj > 0 ? numPasaj : 1);

    const rawCosto = getRowVal(r, ['COSTO TOTAL', 'COSTO', 'COSTO POR VIAJE', 'COSTO BUS', 'COSTO_TOTAL']);
    const costoNum = parseFloat(String(rawCosto || '0').replace(/[^0-9.-]+/g, "")) || 0;
    st.costo += costoNum;
  });

  const rutasArray = Object.keys(statsRuta).map(ruta => {
    const st = statsRuta[ruta];
    const pctOcup = st.capacidad > 0 ? st.pasajeros / st.capacidad : 0;
    const costoViaje = st.viajes > 0 ? st.costo / st.viajes : 0;
    const costoPasaj = st.pasajeros > 0 ? st.costo / st.pasajeros : 0;

    let estado = '-';
    let estadoClass = '';
    if (pctOcup >= 0.9) {
      estado = '✅ Óptimo';
      estadoClass = 'badge badge-success';
    } else if (pctOcup >= 0.7) {
      estado = '🔄 OK';
      estadoClass = 'badge badge-info';
    } else if (pctOcup >= 0.5) {
      estado = '⚠️ Bajo';
      estadoClass = 'badge badge-warning';
    } else {
      estado = '🔴 Muy Bajo';
      estadoClass = 'badge badge-danger';
    }

    return {
      ruta,
      viajes: st.viajes,
      pasajeros: st.pasajeros,
      capacidad: st.capacidad,
      pctOcup,
      costo: st.costo,
      costoViaje,
      costoPasaj,
      estado,
      estadoClass
    };
  }).sort((a, b) => a.ruta.localeCompare(b.ruta, undefined, { numeric: true }));

  // 5. Totales Ejecutivos de la Semana
  let totalViajes = 0, totalPasajeros = 0, totalCapacidad = 0, totalCosto = 0;
  rutasArray.forEach(r => {
    totalViajes += r.viajes;
    totalPasajeros += r.pasajeros;
    totalCapacidad += r.capacidad;
    totalCosto += r.costo;
  });

  const ocupacionPromGlobal = totalCapacidad > 0 ? totalPasajeros / totalCapacidad : 0;
  const costoPromPasajero = totalPasajeros > 0 ? totalCosto / totalPasajeros : 0;
  const costoPromViaje = totalViajes > 0 ? totalCosto / totalViajes : 0;
  const rutasUnicasCount = rutasArray.filter(r => r.viajes > 0).length;

  // Pintar KPIs en pantalla
  const elViajes = document.getElementById('costTotalViajes');
  if (elViajes) elViajes.innerText = totalViajes;

  const elPasaj = document.getElementById('costTotalPasajeros');
  if (elPasaj) elPasaj.innerText = totalPasajeros.toLocaleString('es-PE');

  const elCap = document.getElementById('costCapacidadTotal');
  if (elCap) elCap.innerText = totalCapacidad.toLocaleString('es-PE');

  const elOcup = document.getElementById('costOcupacionProm');
  if (elOcup) elOcup.innerText = `${(ocupacionPromGlobal * 100).toFixed(0)}%`;

  const elCosto = document.getElementById('costTotalSemana');
  if (elCosto) elCosto.innerText = `S/ ${totalCosto.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

  const elCostoPasaj = document.getElementById('costPromPasajero');
  if (elCostoPasaj) elCostoPasaj.innerText = `S/ ${costoPromPasajero.toFixed(2)}`;

  const elCostoViaje = document.getElementById('costPromViaje');
  if (elCostoViaje) elCostoViaje.innerText = `S/ ${costoPromViaje.toFixed(2)}`;

  const elRutasU = document.getElementById('costRutasUnicas');
  if (elRutasU) elRutasU.innerText = rutasUnicasCount;

  // 6. Alertas de Flota
  const rutasBajas = rutasArray.filter(r => r.pctOcup < 0.5).length;
  const rutasAltas = rutasArray.filter(r => r.pctOcup >= 0.9).length;

  let eficienciaGlobalTexto = '🔄 Aceptable';
  let eficienciaColor = '#38bdf8';
  let eficienciaMsg = 'Desempeño general balanceado.';

  if (rutasAltas > 7) {
    eficienciaGlobalTexto = '🟢 Excelente';
    eficienciaColor = '#4ade80';
    eficienciaMsg = 'Alta rentabilidad y demanda sólida.';
  } else if (rutasBajas > 5) {
    eficienciaGlobalTexto = '🔴 Crítico';
    eficienciaColor = '#f87171';
    eficienciaMsg = 'Alta subutilización de buses. Requiere ajuste urgente a unidades menores.';
  } else if (rutasBajas > 2) {
    eficienciaGlobalTexto = '⚠️ Mejorar';
    eficienciaColor = '#fbbf24';
    eficienciaMsg = 'Existen varias rutas con capacidad ociosa por optimizar.';
  }

  const elAltas = document.getElementById('alertRutasAltas');
  if (elAltas) elAltas.innerText = rutasAltas;

  const elBajas = document.getElementById('alertRutasBajas');
  if (elBajas) elBajas.innerText = rutasBajas;

  const elEfic = document.getElementById('alertEficienciaGlobal');
  if (elEfic) {
    elEfic.innerText = eficienciaGlobalTexto;
    elEfic.style.color = eficienciaColor;
  }

  const elEficMsg = document.getElementById('alertEficienciaMsg');
  if (elEficMsg) elEficMsg.innerText = eficienciaMsg;

  // 7. Tabla Detalle por Ruta
  const tbodyDetalle = document.getElementById('tbodyDetalleCostosSemana');
  if (tbodyDetalle) {
    if (rutasArray.length === 0) {
      tbodyDetalle.innerHTML = `<tr><td colspan="9" style="text-align: center; color: #94a3b8; padding: 25px;">No se registraron viajes para la semana ${semSel}.</td></tr>`;
    } else {
      tbodyDetalle.innerHTML = rutasArray.map(r => `
        <tr>
          <td style="font-weight: 600; color: #f8fafc;">${r.ruta}</td>
          <td style="text-align: center;">${r.viajes}</td>
          <td style="text-align: center;">${r.pasajeros}</td>
          <td style="text-align: center;">${r.capacidad}</td>
          <td style="text-align: center; font-weight: 700; color: ${r.pctOcup < 0.5 ? '#f87171' : (r.pctOcup >= 0.7 ? '#4ade80' : '#fbbf24')};">${(r.pctOcup * 100).toFixed(0)}%</td>
          <td style="text-align: right; color: #60a5fa; font-weight: 600;">S/ ${r.costo.toLocaleString('es-PE', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</td>
          <td style="text-align: right;">S/ ${r.costoViaje.toFixed(0)}</td>
          <td style="text-align: right;">S/ ${r.costoPasaj.toFixed(2)}</td>
          <td style="text-align: center;"><span class="${r.estadoClass}">${r.estado}</span></td>
        </tr>
      `).join('');
    }
  }

  const tfootDetalle = document.getElementById('tfootDetalleCostosSemana');
  if (tfootDetalle) {
    tfootDetalle.innerHTML = `
      <tr>
        <td style="color: #60a5fa;">TOTAL</td>
        <td style="text-align: center;">${totalViajes}</td>
        <td style="text-align: center;">${totalPasajeros.toLocaleString('es-PE')}</td>
        <td style="text-align: center;">${totalCapacidad.toLocaleString('es-PE')}</td>
        <td style="text-align: center; color: #38bdf8;">${(ocupacionPromGlobal * 100).toFixed(0)}%</td>
        <td style="text-align: right; color: #60a5fa;">S/ ${totalCosto.toLocaleString('es-PE', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</td>
        <td style="text-align: right;">S/ ${costoPromViaje.toFixed(0)}</td>
        <td style="text-align: right;">S/ ${costoPromPasajero.toFixed(2)}</td>
        <td style="text-align: center;">-</td>
      </tr>
    `;
  }

  // 8. Rankings Top 3 Más y Menos Eficientes
  const validRutas = rutasArray.filter(r => r.viajes > 0 && r.capacidad > 0);
  const top3Mas = [...validRutas].sort((a, b) => b.pctOcup - a.pctOcup).slice(0, 3);
  const top3Menos = [...validRutas].sort((a, b) => a.pctOcup - b.pctOcup).slice(0, 3);

  const tbTop = document.getElementById('tbodyTopEficientes');
  if (tbTop) {
    tbTop.innerHTML = top3Mas.map((r, i) => `
      <tr>
        <td style="font-weight: 600; color: #4ade80;">${r.ruta}</td>
        <td style="text-align: center; font-weight: 700;">${(r.pctOcup * 100).toFixed(0)}%</td>
        <td style="text-align: right; color: #cbd5e1;">S/ ${r.costoViaje.toFixed(0)}</td>
      </tr>
    `).join('');
  }

  const tbBottom = document.getElementById('tbodyBottomEficientes');
  if (tbBottom) {
    tbBottom.innerHTML = top3Menos.map((r, i) => `
      <tr>
        <td style="font-weight: 600; color: #f87171;">${r.ruta}</td>
        <td style="text-align: center; font-weight: 700; color: #f87171;">${(r.pctOcup * 100).toFixed(0)}%</td>
        <td style="text-align: right; color: #cbd5e1;">S/ ${r.costoViaje.toFixed(0)}</td>
      </tr>
    `).join('');
  }

  // 9. Recomendaciones de Toma de Decisiones
  const p1Urgente = rutasBajas > 5 ? 'prioridad-urgente' : (rutasBajas > 2 ? 'prioridad-media' : 'prioridad-baja');
  const p1Texto = rutasBajas > 5 ? '🔴 URGENTE' : (rutasBajas > 2 ? '🟡 MEDIA' : '🟢 BAJA');

  const p2Urgente = rutasAltas > 3 ? 'prioridad-urgente' : (rutasAltas > 1 ? 'prioridad-media' : 'prioridad-baja');
  const p2Texto = rutasAltas > 3 ? '🔴 URGENTE' : (rutasAltas > 1 ? '🟡 MEDIA' : '🟢 BAJA');

  const p3Urgente = costoPromPasajero > 20 ? 'prioridad-urgente' : (costoPromPasajero > 15 ? 'prioridad-media' : 'prioridad-baja');
  const p3Texto = costoPromPasajero > 20 ? '🔴 URGENTE' : (costoPromPasajero > 15 ? '🟡 MEDIA' : '🟢 BAJA');
  const p3Recom = costoPromPasajero > 15 ? '⚠️ Optimizar rutas para reducir costo/pasajero' : '✅ Costo por pasajero controlado y aceptable';

  const p4Urgente = ocupacionPromGlobal < 0.5 ? 'prioridad-urgente' : (ocupacionPromGlobal < 0.7 ? 'prioridad-media' : 'prioridad-baja');
  const p4Texto = ocupacionPromGlobal < 0.5 ? '🔴 URGENTE' : (ocupacionPromGlobal < 0.7 ? '🟡 MEDIA' : '🟢 BAJA');
  const p4Recom = ocupacionPromGlobal < 0.6 ? '⚠️ Revisar dimensionamiento general de flota' : (ocupacionPromGlobal < 0.75 ? '🔄 Monitorear tendencia operativa' : '✅ Ocupación saludable');

  const p5Urgente = totalCosto > 35000 ? 'prioridad-urgente' : (totalCosto > 30000 ? 'prioridad-media' : 'prioridad-baja');
  const p5Texto = totalCosto > 35000 ? '🔴 URGENTE' : (totalCosto > 30000 ? '🟡 MEDIA' : '🟢 BAJA');
  const p5Recom = totalCosto > 30000 ? '📊 Evaluar renegociación con proveedores' : '✅ Dentro de presupuesto esperado';

  const tbRecom = document.getElementById('tbodyRecomendacionesDecisiones');
  if (tbRecom) {
    tbRecom.innerHTML = `
      <tr>
        <td style="font-weight: 700; color: #94a3b8;">1</td>
        <td style="font-weight: 600; color: #f8fafc;">Rutas con muy baja ocupación</td>
        <td>${rutasBajas} rutas &lt; 50%</td>
        <td>Considerar cambio a vans o vehículos menores (Sprinter/Custer)</td>
        <td style="text-align: center;"><span class="prioridad-pill ${p1Urgente}">${p1Texto}</span></td>
        <td style="color: #4ade80; font-weight: 600;">Ahorro 20-40% en costos</td>
      </tr>
      <tr>
        <td style="font-weight: 700; color: #94a3b8;">2</td>
        <td style="font-weight: 600; color: #f8fafc;">Rutas con alta demanda</td>
        <td>${rutasAltas} rutas ≥ 90%</td>
        <td>Considerar bus adicional si supera capacidad asignada</td>
        <td style="text-align: center;"><span class="prioridad-pill ${p2Urgente}">${p2Texto}</span></td>
        <td style="color: #60a5fa;">Mejora satisfacción colaboradores</td>
      </tr>
      <tr>
        <td style="font-weight: 700; color: #94a3b8;">3</td>
        <td style="font-weight: 600; color: #f8fafc;">Costo promedio por pasajero</td>
        <td>S/ ${costoPromPasajero.toFixed(2)}</td>
        <td>${p3Recom}</td>
        <td style="text-align: center;"><span class="prioridad-pill ${p3Urgente}">${p3Texto}</span></td>
        <td style="color: #cbd5e1;">Meta: &lt; S/ 15.00</td>
      </tr>
      <tr>
        <td style="font-weight: 700; color: #94a3b8;">4</td>
        <td style="font-weight: 600; color: #f8fafc;">Ocupación promedio global</td>
        <td>${(ocupacionPromGlobal * 100).toFixed(0)}%</td>
        <td>${p4Recom}</td>
        <td style="text-align: center;"><span class="prioridad-pill ${p4Urgente}">${p4Texto}</span></td>
        <td style="color: #cbd5e1;">Meta: &gt; 70%</td>
      </tr>
      <tr>
        <td style="font-weight: 700; color: #94a3b8;">5</td>
        <td style="font-weight: 600; color: #f8fafc;">Inversión semanal en transporte</td>
        <td>S/ ${Math.round(totalCosto).toLocaleString('es-PE')}</td>
        <td>${p5Recom}</td>
        <td style="text-align: center;"><span class="prioridad-pill ${p5Urgente}">${p5Texto}</span></td>
        <td style="color: #cbd5e1;">Benchmark: S/ 30,000/sem</td>
      </tr>
    `;
  }

  // 10. Proyección Mensual (x4)
  const tbProy = document.getElementById('tbodyProyeccionMensual');
  if (tbProy) {
    tbProy.innerHTML = `
      <tr>
        <td style="font-weight: 600; color: #f8fafc;">Costo Total</td>
        <td style="text-align: right; color: #60a5fa; font-weight: 600;">S/ ${Math.round(totalCosto).toLocaleString('es-PE')}</td>
        <td style="text-align: right; color: #38bdf8; font-weight: 700;">S/ ${Math.round(totalCosto * 4).toLocaleString('es-PE')}</td>
      </tr>
      <tr>
        <td style="font-weight: 600; color: #f8fafc;">Pasajeros</td>
        <td style="text-align: right;">${totalPasajeros.toLocaleString('es-PE')}</td>
        <td style="text-align: right; font-weight: 600;">${(totalPasajeros * 4).toLocaleString('es-PE')}</td>
      </tr>
      <tr>
        <td style="font-weight: 600; color: #f8fafc;">Viajes</td>
        <td style="text-align: right;">${totalViajes}</td>
        <td style="text-align: right; font-weight: 600;">${totalViajes * 4}</td>
      </tr>
      <tr>
        <td style="font-weight: 600; color: #f8fafc;">Costo/Pasajero</td>
        <td style="text-align: right;">S/ ${costoPromPasajero.toFixed(2)}</td>
        <td style="text-align: right; font-weight: 600;">S/ ${(costoPromPasajero * 4).toFixed(2)}</td>
      </tr>
    `;
  }

  // 11. Gráficos de Costos (Solo si la pestaña está visible para evitar problemas de canvas oculto)
  const viewCostosEl = document.getElementById('viewAnalisisCostos');
  if (viewCostosEl && !viewCostosEl.classList.contains('hidden')) {
    renderGraficosCostos(rutasArray);
  }
}

// =========================================================
// Gráficos de Costos y Eficiencia (Chart.js)
// =========================================================

function renderGraficosCostos(rutasArray) {
  // Gráfico 1: % Ocupación por Ruta
  const ctxOcup = document.getElementById('chartOcupacionRutasCostos');
  if (ctxOcup) {
    const labels = rutasArray.map(r => r.ruta.replace('RUTA ', 'R.'));
    const dataOcup = rutasArray.map(r => Math.round(r.pctOcup * 100));
    const colorsOcup = rutasArray.map(r => {
      if (r.pctOcup >= 0.9) return '#22c55e'; // verde
      if (r.pctOcup >= 0.7) return '#3b82f6'; // azul
      if (r.pctOcup >= 0.5) return '#eab308'; // amarillo
      return '#ef4444'; // rojo
    });

    if (AppState.charts['chartOcupacionRutasCostos']) {
      AppState.charts['chartOcupacionRutasCostos'].destroy();
    }

    AppState.charts['chartOcupacionRutasCostos'] = new Chart(ctxOcup, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: '% Ocupación',
          data: dataOcup,
          backgroundColor: colorsOcup,
          borderRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            ticks: {
              callback: (v) => v + '%'
            },
            grid: { color: 'rgba(255, 255, 255, 0.08)' }
          },
          x: {
            grid: { display: false }
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => ` ${ctx.parsed.y}% de ocupación`
            }
          }
        }
      }
    });
  }

  // Gráfico 2: Costo Total vs Pasajeros por Ruta
  const ctxCostoPasaj = document.getElementById('chartCostoPasajerosRutas');
  if (ctxCostoPasaj) {
    const labels = rutasArray.map(r => r.ruta.replace('RUTA ', 'R.'));
    const dataCosto = rutasArray.map(r => r.costo);
    const dataPasaj = rutasArray.map(r => r.pasajeros);

    if (AppState.charts['chartCostoPasajerosRutas']) {
      AppState.charts['chartCostoPasajerosRutas'].destroy();
    }

    AppState.charts['chartCostoPasajerosRutas'] = new Chart(ctxCostoPasaj, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'Costo Total (S/)',
            data: dataCosto,
            backgroundColor: '#3b82f6',
            borderRadius: 6,
            yAxisID: 'y'
          },
          {
            label: 'Pasajeros Atendidos',
            data: dataPasaj,
            backgroundColor: '#10b981',
            borderRadius: 6,
            yAxisID: 'y1'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            type: 'linear',
            position: 'left',
            ticks: {
              callback: (v) => 'S/ ' + v
            },
            grid: { color: 'rgba(255, 255, 255, 0.08)' }
          },
          y1: {
            type: 'linear',
            position: 'right',
            grid: { drawOnChartArea: false },
            ticks: { precision: 0 }
          },
          x: { grid: { display: false } }
        },
        plugins: {
          legend: {
            position: 'top',
            labels: { color: '#cbd5e1', boxWidth: 12 }
          }
        }
      }
    });
  }
}
