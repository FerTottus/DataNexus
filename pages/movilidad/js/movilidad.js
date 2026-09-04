/**
 * Script Principal para el Dashboard de Movilidad
 */

// Estado global de la aplicación
const AppState = {
  sheetId: '',
  sheetUrl: '',
  rawEmployees: [], // Todos los empleados combinados
  filteredEmployees: [], // Empleados después de aplicar filtros
  charts: {} // Referencias a instancias de Chart.js
};

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

  // Chip Group Logic para multi-selección
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
    
    const excelWarningHtml = isExcel ? `<span class="file-badge custom-tooltip-icon" style="background: #fee2e2; color: #ef4444; border: 1px solid #fecaca; margin-right: 5px;" data-tooltip="Ábrelo en Drive y dale a 'Guardar como Hoja de cálculo de Google'">⚠️ Inválido (.xlsx)</span>` : '';

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
  const term = searchTerm.toLowerCase();
  const items = document.querySelectorAll('#drive-file-list .file-list-item');
  let hasVisible = false;
  
  items.forEach(item => {
    if (item.dataset.filename.includes(term)) {
      item.style.display = 'flex';
      hasVisible = true;
    } else {
      item.style.display = 'none';
    }
  });
  
  const empty = document.getElementById('drive-empty');
  if (hasVisible) {
    empty.classList.add('hidden');
  } else {
    empty.classList.remove('hidden');
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
    userInfo.classList.remove('hidden');
    userEmailSpan.innerText = GoogleSheetsService.userEmail || 'Conectado';
  } else {
    btnLogin.classList.remove('hidden');
    userInfo.classList.add('hidden');
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
    let registroData = [];

    const processSheet = (sheetData, areaName) => {
      if (!sheetData) return;
      sheetData.rows.forEach(row => {
        // Ignorar filas vacías o rotulos de información sin DNI o Nombres reales
        if(!row['Distrito']) return;
        
        combined.push({
          dni: String(row['DNI'] || row['DOCUMENTO'] || row['ID'] || '').trim(),
          area: areaName,
          tipo: String(row['CLASIFICACION'] || '').toUpperCase().trim(),
          distrito: String(row['Distrito'] || '').trim(),
          distCd: parseFloat(row['DIST. AL CD (km)']) || 0,
          clasifCd: String(row['CLASIF. DIST. CD'] || '').trim(),
          distParadero: parseFloat(row['DIST. PARADERO (km)']) || 0,
          clasifParadero: String(row['CLASIF. DIST. PARADERO'] || '').trim(),
          ruta: String(row['RUTA PARADERO'] || '').trim(),
          paradero: String(row['PARADERO'] || row['PARADERO MÁS CERCANO'] || row['NOMBRE PARADERO'] || '').trim()
        });
      });
    };

    processSheet(secos, 'SECOS');
    processSheet(ppa, 'PPA');
    processSheet(frescos, 'FRESCOS');

    AppState.rawEmployees = combined;
    AppState.filteredEmployees = [...combined];
    
    // Guardamos las filas raw del registro diario para agregarlas dinámicamente según filtros
    AppState.rawRegistroDiario = (registroDiario && registroDiario.rows) ? registroDiario.rows : [];

    // Poblar dropdown de fechas basado en los datos únicos
    const fSet = new Set();
    AppState.rawRegistroDiario.forEach(r => {
      let f = r['FECHA'] || r['FECHA DE VIAJE'] || '';
      f = String(f).split(' ')[0];
      if(f) fSet.add(f);
    });
    const fechas = Array.from(fSet).sort((a,b) => {
      // Orden simple de string, idealmente DD/MM/YYYY
      return a.localeCompare(b);
    });
    
    const selFecha = document.getElementById('filtroFecha');
    if (selFecha) {
      selFecha.innerHTML = `<option value="ULTIMA">Última Fecha Disponible</option><option value="TODAS">Ver Todas</option>`;
      fechas.forEach(f => {
        selFecha.innerHTML += `<option value="${f}">${f}</option>`;
      });
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

  // Filtramos la data demográfica de empleados
  AppState.filteredEmployees = AppState.rawEmployees.filter(emp => {
    const matchArea = areas.includes('TODOS') || areas.includes(emp.area);
    const matchTipo = tipos.includes('TODOS') || tipos.includes(emp.tipo);
    return matchArea && matchTipo;
  });

  // Filtramos y agregamos la data del registro diario de buses
  const empMap = new Map();
  AppState.filteredEmployees.forEach(e => {
    if (e.dni) empMap.set(e.dni, true);
  });

  const selF = document.getElementById('filtroFecha');
  const valFecha = selF ? selF.value : 'ULTIMA';
  const selD = document.getElementById('filtroDia');
  const valDia = selD ? selD.value : 'TODOS';

  const aggrMap = {};
  AppState.rawRegistroDiario.forEach(row => {
    const getVal = (keys) => {
      for (let k of keys) {
        if (row[k] !== undefined && row[k] !== '') return row[k];
      }
      return null;
    };

    const dni = String(getVal(['DNI', 'DOCUMENTO', 'ID']) || '').trim();
    // CRÍTICO: Si el empleado no está en el mapa de empleados filtrados, LO IGNORAMOS.
    // Solo contamos a los empleados que cumplen el filtro de Área y Tipo
    if (dni && !empMap.has(dni)) return;

    let rawFecha = getVal(['FECHA', 'FECHA DE VIAJE']) || '';
    const fechaSoloDia = String(rawFecha).split(' ')[0]; 

    // Filtros de fecha y día
    const diaVal = getVal(['DÍA', 'DIA']);
    if (valDia !== 'TODOS' && diaVal && String(diaVal).toLowerCase() !== valDia.toLowerCase()) return;
    if (valFecha !== 'ULTIMA' && valFecha !== 'TODAS' && fechaSoloDia !== valFecha) return;

    const rutaVal = getVal(['RUTA', 'RUTA ASIGNADA']);
    if (!rutaVal) return;

    const key = `${fechaSoloDia}|${rutaVal}`;
    if (!aggrMap[key]) {
      const rawCosto = getVal(['COSTO TOTAL', 'COSTO', 'COSTO POR VIAJE', 'COSTO BUS']);
      const costoNum = parseFloat(String(rawCosto || '0').replace(/[^0-9.-]+/g, "")) || 0;

      aggrMap[key] = {
        dia: diaVal,
        fecha: fechaSoloDia,
        semana: parseInt(getVal(['SEMANA'])) || 0,
        ruta: rutaVal,
        capacidad: parseFloat(getVal(['CAPACIDAD', 'CAPACIDAD DE BUS', 'CAPACIDAD BUS'])) || 0,
        pasajeros: 0,
        costo: costoNum 
      };
    }

    aggrMap[key].pasajeros += 1;
  });

  AppState.registroData = Object.values(aggrMap);
  AppState.filtroFechaModo = valFecha; // Guardamos para usarlo luego
  
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
  if (regData && regData.length > 0) {
    document.getElementById('seccionRegistroDiario').style.display = 'block';
    document.getElementById('seccionDashboardBuses').style.display = 'block';

    // En lugar de confiar en que los números de SEMANA siempre sean ascendentes 
    // (ya que a veces en Excel se reinician o calculan mal, ej. de 36 a 21),
    // simplemente tomamos la última fila registrada en el archivo.
    const ultimaFila = regData[regData.length - 1];
    const semanaActual = ultimaFila.semana;
    const diaActual = ultimaFila.dia;
    const fechaActual = ultimaFila.fecha;

    // Filtro por semana y día actual basados en esa última fila
    const datosSemana = regData.filter(r => r.semana === semanaActual);
    
    // Si tenemos la fecha exacta (ej. 03/09/2026), filtramos por eso para evitar 
    // colisiones con otros jueves de la misma semana (o semanas mal numeradas)
    const datosDia = fechaActual 
      ? regData.filter(r => r.fecha === fechaActual)
      : datosSemana.filter(r => r.dia === diaActual);

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

    // Dashboard Día / Semana
    const aggDia = { viajes: 0, pasaj: 0, costo: 0, cap: 0 };
    datosDia.forEach(r => { aggDia.viajes++; aggDia.pasaj += r.pasajeros; aggDia.costo += r.costo; aggDia.cap += r.capacidad; });
    
    const aggSem = { viajes: 0, pasaj: 0, costo: 0, cap: 0 };
    datosSemana.forEach(r => { aggSem.viajes++; aggSem.pasaj += r.pasajeros; aggSem.costo += r.costo; aggSem.cap += r.capacidad; });

    // Actualizar KPIs de la parte superior
    document.getElementById('kpiCostoDia').innerText = 'S/ ' + aggDia.costo.toFixed(2);
    document.getElementById('kpiOcupacionDia').innerText = formatPct(aggDia.pasaj, aggDia.cap);

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
    const s1 = document.getElementById('seccionRegistroDiario');
    const s2 = document.getElementById('seccionDashboardBuses');
    if(s1) s1.style.display = 'none';
    if(s2) s2.style.display = 'none';
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
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom' }
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

