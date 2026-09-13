/**
 * DataNexus - Antigüedad de Inventario (Secos 655 y Frescos 676)
 * Motor Ultra-Optimizado de Procesamiento por Lotes y Captura 16:9 Nativa
 * 
 * 1. Procesamiento ultra-rápido por lotes (5,000-8,000 filas/chunk) sin congelar la UI.
 * 2. Suma aritmética exacta de LPNs en todos los niveles.
 * 3. Gráfico de Barras con valores y porcentajes ENCIMA de cada barra.
 * 4. Tabla de resumen de rangos (Table 1) en Slide 1 sincronizada con tbBD.
 * 5. Captura 16:9 panorámica (1920x1080) directa al portapapeles (Ctrl + V en PPT) sin descargas forzadas.
 */

// Estado global
let currentWarehouse = '655'; // '655' (Secos) o '676' (Frescos)
let currentSlide = 1;
let currentWeekLabel = 'Semana 37';
let currentChartS1Mode = 'bultos'; // 'bultos' o 'costo'
let chartInstanceS1 = null;
let chartInstanceS3 = null;
let isCapturing = false;
let lastCapturedBlob = null;
let lastCapturedDataUrl = null;

// Base de datos activa
let ACTIVE_DATABASE = {
  '655': null,
  '676': null
};

// ══════════════════════════════════════════════════════════════════════════════
// 1. UTILIDADES Y FORMATEADORES
// ══════════════════════════════════════════════════════════════════════════════

function safeFixed(val, decimals = 1, suffix = '') {
  if (val === null || val === undefined || isNaN(Number(val))) {
    return '0.0' + suffix;
  }
  return Number(val).toFixed(decimals) + suffix;
}

function parseNum(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  if (!v) return 0;
  let str = String(v).trim().replace(/\s+/g, '');
  if (str.includes(',') && str.includes('.')) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.includes(',')) {
    const parts = str.split(',');
    if (parts.length === 2 && parts[1].length !== 3) {
      str = str.replace(',', '.');
    } else if (parts.length > 2) {
      str = str.replace(/,/g, '');
    } else {
      str = str.replace(/,/g, '');
    }
  }
  const clean = str.replace(/[^\d\.-]/g, '').trim();
  const n = parseFloat(clean);
  return isNaN(n) ? 0 : n;
}

function formatCurrency(val) {
  const n = parseNum(val);
  return 'S/ ' + Math.round(n).toLocaleString('en-US');
}

function formatNumber(val) {
  const n = parseNum(val);
  return Math.round(n).toLocaleString('en-US');
}

function formatCompact(val) {
  const n = parseNum(val);
  if (!n) return '0';
  if (n >= 1000000) return (n / 1000000).toFixed(1) + 'M';
  if (n >= 1000) return (n / 1000).toFixed(0) + 'K';
  return Math.round(n).toString();
}

function showToast(message, type = 'info', duration = 3500) {
  const toast = document.getElementById('toastMessage');
  const toastText = document.getElementById('toastText');
  if (!toast || !toastText) return;

  toastText.textContent = message;
  toast.style.background = type === 'success' ? '#065f46' : (type === 'danger' ? '#991b1b' : '#0f172a');
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), duration);
}

function showLoading(text, subtext = '') {
  const overlay = document.getElementById('loadingOverlay');
  const txt = document.getElementById('loadingText');
  const sub = document.getElementById('loadingSubtext');
  if (txt) txt.textContent = text || 'Procesando archivo Excel...';
  if (sub) sub.textContent = subtext;
  if (overlay) overlay.style.display = 'flex';
}

function hideLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) overlay.style.display = 'none';
}

function yieldToUi() {
  return new Promise(resolve => setTimeout(resolve, 0));
}

function normalizeRange(str) {
  if (!str) return '0 a 10 Semanas';
  const s = String(str).toLowerCase().trim();
  if (s.includes('0 a 10') || s.includes('0-10') || s.includes('< 10') || s.includes('menor a 10')) return '0 a 10 Semanas';
  if (s.includes('10 a 25') || s.includes('10-25')) return '10 a 25 Semanas';
  if (s.includes('25 a 52') || s.includes('25-52')) return '25 a 52 Semanas';
  if (s.includes('52') || s.includes('año') || s.includes('ano') || s.includes('> 52') || s.includes('mayor')) return 'mayor a 52 Semanas';
  return '0 a 10 Semanas';
}

// ══════════════════════════════════════════════════════════════════════════════
// 2. DETECTOR Y MAPA DE ÍNDICES DE COLUMNAS (ALTO RENDIMIENTO O(1))
// ══════════════════════════════════════════════════════════════════════════════

function buildColumnIndexMap(headers) {
  const normHeaders = headers.map(h => String(h || '').toLowerCase().replace(/[\s_\-\.\:\/%]/g, '').trim());

  function findIdx(candidates) {
    // 1. Coincidencia exacta limpia
    for (const cand of candidates) {
      const cleanCand = cand.toLowerCase().replace(/[\s_\-\.\:\/%]/g, '').trim();
      for (let i = 0; i < normHeaders.length; i++) {
        if (normHeaders[i] === cleanCand) return i;
      }
    }
    // 2. Coincidencia parcial si no hubo exacta
    for (const cand of candidates) {
      const cleanCand = cand.toLowerCase().replace(/[\s_\-\.\:\/%]/g, '').trim();
      for (let i = 0; i < normHeaders.length; i++) {
        if (normHeaders[i].includes(cleanCand)) return i;
      }
    }
    return -1;
  }

  return {
    whse: findIdx(['whse', 'almacenfisico', 'almacen', 'almacen_fisico']),
    costos: findIdx(['costos', 'costo', 'costototal', 'costo_total']),
    bultos: findIdx(['bultos', 'bulto', 'cajas', 'cantidad']),
    lpn: findIdx(['lpn', 'nrolpn', 'lpns', 'cantidadlpn', 'cantlpn']),
    sku: findIdx(['sku', 'codigo', 'codsku']),
    desc: findIdx(['descripcion', 'desc', 'descripcionsku', 'producto']),
    div: findIdx(['division', 'div', 'dpto', 'departamento']),
    semanas: findIdx(['semanas', 'rangosemanas', 'antiguedadsemanas', 'antiguedad']),
    zona: findIdx(['zona', 'tipozona', 'tipo_zona']),
    ubic: findIdx(['ubicacion', 'posicion', 'slot']),
    rngFv: findIdx(['rngfv', 'fechavencim', 'vencimiento', 'fechavencimiento']),
    onHand: findIdx(['onhand', 'unidades', 'on_hand']),
    dias: findIdx(['dias', 'diasantiguedad', 'antiguedaddias']),
    semana: findIdx(['semana', 'sem', 'week', 'nrosemana'])
  };
}

function findHeaderRowInMatrix(matrix) {
  const maxScan = Math.min(matrix.length, 12);
  for (let r = 0; r < maxScan; r++) {
    const row = matrix[r] || [];
    const text = row.map(c => String(c || '').toUpperCase()).join(' ');
    if ((text.includes('WHSE') || text.includes('ALMACEN')) &&
        (text.includes('SKU') || text.includes('COSTO') || text.includes('LPN') || text.includes('BULT'))) {
      return r;
    }
  }
  return 0;
}

// ══════════════════════════════════════════════════════════════════════════════
// 3. CARGA DE ARCHIVO Y DRAG & DROP
// ══════════════════════════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  setupDropzone();
});

function setupDropzone() {
  const dropzone = document.getElementById('dropzoneBox');
  if (!dropzone) return;

  ['dragenter', 'dragover'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length > 0) {
      processExcelFile(dt.files[0]);
    }
  }, false);
}

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (file) {
    processExcelFile(file);
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// 4. MOTOR PRINCIPAL DE LECTURA DE EXCEL (SHEETJS POR LOTES ASÍNCRONOS)
// ══════════════════════════════════════════════════════════════════════════════

function processExcelFile(file) {
  showLoading(`Abriendo ${file.name}...`, 'Leyendo archivo binario...');

  const reader = new FileReader();

  reader.onload = async function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      showLoading('Inspeccionando hojas del libro...', 'Modo de alto rendimiento activo');
      await yieldToUi();

      // Lectura rápida de SheetJS sin strings pesados
      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: false,
        cellNF: false,
        cellText: false
      });

      // 1. Detectar Hoja BD
      let wsBD = null;
      let sheetNameBD = '';
      for (const sName of workbook.SheetNames) {
        const clean = sName.trim().toUpperCase();
        if (clean === 'BD' || clean === 'TBBD' || clean.includes('BD') || clean.includes('BASE')) {
          wsBD = workbook.Sheets[sName];
          sheetNameBD = sName;
          break;
        }
      }
      if (!wsBD) {
        sheetNameBD = workbook.SheetNames[0];
        wsBD = workbook.Sheets[sheetNameBD];
      }

      showLoading(`Estructurando matriz de datos de "${sheetNameBD}"...`);
      await yieldToUi();

      // Matriz 2D rápida: O(1) de memoria y sin miles de objetos
      const rawMatrix = XLSX.utils.sheet_to_json(wsBD, {
        header: 1,
        defval: ''
      });

      if (!rawMatrix || rawMatrix.length === 0) {
        throw new Error(`La hoja "${sheetNameBD}" no contiene registros.`);
      }

      const headerRowIdx = findHeaderRowInMatrix(rawMatrix);
      const headers = (rawMatrix[headerRowIdx] || []).map(h => String(h || '').trim());
      const colMap = buildColumnIndexMap(headers);

      // 2. Detectar Hoja EVOLUTIVO
      let evolSecos = null;
      let evolFrescos = null;
      const sheetNameEvol = workbook.SheetNames.find(n => {
        const clean = n.trim().toUpperCase();
        return clean.includes('EVOL') || clean.includes('HIST');
      });

      if (sheetNameEvol && workbook.Sheets[sheetNameEvol]) {
        showLoading('Extrayendo evolución de semanas desde hoja EVOLUTIVO...');
        await yieldToUi();
        const wsEvol = workbook.Sheets[sheetNameEvol];
        const evolFound = parseEvolutivoSheetSmart(wsEvol);
        evolSecos = evolFound.secos;
        evolFrescos = evolFound.frescos;
      }

      // 3. Detectar semana actual de los datos
      let activeWeekNum = '37';
      const firstDataRow = rawMatrix[headerRowIdx + 1] || [];
      if (colMap.semana !== -1 && firstDataRow[colMap.semana] !== undefined) {
        const cleanW = String(firstDataRow[colMap.semana]).replace(/[^\d]/g, '').trim();
        if (cleanW) activeWeekNum = cleanW;
      } else if (evolSecos && evolSecos.weeks && evolSecos.weeks.length > 0) {
        const lastW = evolSecos.weeks[evolSecos.weeks.length - 1].replace(/[^\d]/g, '').trim();
        if (lastW) activeWeekNum = lastW;
      }
      currentWeekLabel = `Semana ${activeWeekNum}`;
      const activeWeekTag = 'S-' + activeWeekNum;

      // 4. Procesar filas por bloques asíncronos para evitar congelamientos
      showLoading('Procesando inventario en lotes optimizados...', 'Calculando Secos (655) y Frescos (676)...');
      await yieldToUi();

      const dataRows = rawMatrix.slice(headerRowIdx + 1);
      const compiledData = await compileDataOptimized(dataRows, colMap, evolSecos, evolFrescos, activeWeekTag);

      ACTIVE_DATABASE['655'] = compiledData.secos;
      ACTIVE_DATABASE['676'] = compiledData.frescos;

      // 5. Actualizar interfaz
      document.getElementById('dropzoneBox').style.display = 'none';
      document.getElementById('fileStatusBar').style.display = 'flex';
      document.getElementById('loadedFileName').textContent = `Archivo: ${file.name}`;
      document.getElementById('loadedFileMeta').textContent = `${dataRows.length.toLocaleString('en-US')} filas procesadas | CD Secos 655 y Frescos 676 actualizados`;

      updateHeaderWeekBadges();
      renderWarehouseData(currentWarehouse);
      hideLoading();
      showToast(`¡${dataRows.length.toLocaleString('en-US')} filas procesadas con éxito!`, 'success', 4000);

    } catch (err) {
      console.error('Error al procesar Excel:', err);
      hideLoading();
      alert(`Error al procesar el archivo Excel: ${err.message}`);
    }
  };

  reader.onerror = function () {
    hideLoading();
    showToast('Error al leer el archivo desde el disco', 'danger');
  };

  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════════════════════════════════════════
// 5. EXTRACCIÓN DE LA HOJA EVOLUTIVO
// ══════════════════════════════════════════════════════════════════════════════

function parseEvolutivoSheetSmart(ws) {
  const result = { secos: null, frescos: null };
  if (!ws || !ws['!ref']) return result;

  const range = XLSX.utils.decode_range(ws['!ref']);
  const matchingRows = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= Math.min(range.e.c, 6); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined) {
        const txt = String(cell.v).toLowerCase();
        if (txt.includes('0 a 10') || txt.includes('0-10')) {
          matchingRows.push({ rowDataStart: r, colLabel: c });
          break;
        }
      }
    }
  }

  if (matchingRows.length > 0) {
    result.secos = extractEvolutivoBlock(ws, matchingRows[0].rowDataStart, range);
  } else {
    result.secos = extractEvolutivoBlock(ws, 5, range);
  }

  if (matchingRows.length > 1) {
    result.frescos = extractEvolutivoBlock(ws, matchingRows[1].rowDataStart, range);
  } else {
    result.frescos = extractEvolutivoBlock(ws, 14, range);
  }

  return result;
}

function extractEvolutivoBlock(ws, dataStartRow, range) {
  const headerRow = Math.max(0, dataStartRow - 1);
  const validCols = [];

  for (let c = 1; c <= range.e.c; c++) {
    const hCell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    const dCell = ws[XLSX.utils.encode_cell({ r: dataStartRow, c })];

    if (hCell && hCell.v !== undefined && String(hCell.v).trim() !== '' &&
        dCell && dCell.v !== undefined && String(dCell.v).trim() !== '') {
      validCols.push(c);
    }
  }

  if (validCols.length === 0) return null;

  const lastCols = validCols.slice(-7);

  const weeks = lastCols.map(c => {
    const cell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    let txt = String(cell ? cell.v : '').trim();
    if (!txt.toUpperCase().startsWith('S-') && !txt.toUpperCase().startsWith('SEM')) {
      txt = 'S-' + txt;
    }
    return txt;
  });

  const rangeDefs = [
    { label: '0 a 10 Semanas', color: '#10b981' },
    { label: '10 a 25 Semanas', color: '#f59e0b' },
    { label: '25 a 52 Semanas', color: '#f97316' },
    { label: 'Mayor a 52 Semanas', color: '#ef4444' }
  ];

  const evolution = rangeDefs.map((rDef, idx) => {
    const r = dataStartRow + idx;
    const values = lastCols.map(c => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v === undefined) return 0;
      
      let val = 0;
      if (typeof cell.v === 'number') {
        val = cell.v <= 1.0 && cell.v > 0 ? cell.v * 100 : cell.v;
      } else if (cell.w && typeof cell.w === 'string' && cell.w.includes('%')) {
        val = parseFloat(cell.w.replace('%', '').replace(',', '.').trim()) || 0;
      } else {
        const clean = String(cell.v || '').replace('%', '').replace(',', '.').trim();
        val = parseFloat(clean) || 0;
        if (val <= 1.0 && val > 0) val = val * 100;
      }
      return isNaN(val) ? 0 : Number(val.toFixed(5));
    });
    return {
      label: rDef.label,
      values,
      color: rDef.color
    };
  });

  return { weeks, evolution };
}

// ══════════════════════════════════════════════════════════════════════════════
// 6. MOTOR POR LOTES ASÍNCRONOS PARA EVITAR CONGELAMIENTOS
// ══════════════════════════════════════════════════════════════════════════════

function initAccumulator(whseCode, whseLabel) {
  return {
    whseCode,
    whseLabel,
    totalCost: 0,
    totalBultos: 0,
    totalLpns: 0,
    rangeAgg: {
      '0 a 10 Semanas': { cost: 0, bultos: 0, lpns: 0, color: '#10b981', status: 'Saludable' },
      '10 a 25 Semanas': { cost: 0, bultos: 0, lpns: 0, color: '#f59e0b', status: 'En Alerta' },
      '25 a 52 Semanas': { cost: 0, bultos: 0, lpns: 0, color: '#f97316', status: 'Riesgo Medio' },
      'mayor a 52 Semanas': { cost: 0, bultos: 0, lpns: 0, color: '#ef4444', status: 'Crítico >1 año' }
    },
    divisionAgg: {},
    zoneAgg: {
      rck: { totalCost: 0, lpns: 0, bultos: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, skus52: {} },
      rhb: { totalCost: 0, lpns: 0, bultos: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, skus52: {} }
    },
    skusOver52: {},
    locationsAgg: {}
  };
}

function accumulateRow(row, colMap, acc) {
  const cost = colMap.costos !== -1 ? parseNum(row[colMap.costos]) : 0;
  const bultos = colMap.bultos !== -1 ? parseNum(row[colMap.bultos]) : 0;
  const onHand = colMap.onHand !== -1 ? parseNum(row[colMap.onHand]) : 0;
  const dias = colMap.dias !== -1 ? parseNum(row[colMap.dias]) : 0;
  const lpnVal = 1; // Cada registro en tbBD representa exactamente 1 pallet / 1 LPN físico

  const sku = colMap.sku !== -1 ? String(row[colMap.sku] || '').trim() : '';
  const desc = colMap.desc !== -1 ? String(row[colMap.desc] || '').trim() : '';
  const div = colMap.div !== -1 ? (String(row[colMap.div] || 'OTROS').trim() || 'OTROS') : 'OTROS';
  const rng = colMap.semanas !== -1 ? normalizeRange(row[colMap.semanas]) : '0 a 10 Semanas';
  const zona = colMap.zona !== -1 ? String(row[colMap.zona] || '').toUpperCase().trim() : '';
  const ubic = colMap.ubic !== -1 ? String(row[colMap.ubic] || '').trim() : '';
  const rngFv = colMap.rngFv !== -1 ? (String(row[colMap.rngFv] || '-').trim() || '-') : '-';

  acc.totalCost += cost;
  acc.totalBultos += bultos;
  acc.totalLpns += lpnVal;

  // 1. Rangos
  if (acc.rangeAgg[rng]) {
    acc.rangeAgg[rng].cost += cost;
    acc.rangeAgg[rng].bultos += bultos;
    acc.rangeAgg[rng].lpns += lpnVal;
  }

  // 2. Divisiones
  if (!acc.divisionAgg[div]) {
    acc.divisionAgg[div] = { code: div, r010: 0, r1025: 0, r2552: 0, r52: 0, total: 0 };
  }
  acc.divisionAgg[div].total += cost;
  if (rng === '0 a 10 Semanas') acc.divisionAgg[div].r010 += cost;
  else if (rng === '10 a 25 Semanas') acc.divisionAgg[div].r1025 += cost;
  else if (rng === '25 a 52 Semanas') acc.divisionAgg[div].r2552 += cost;
  else if (rng === 'mayor a 52 Semanas') acc.divisionAgg[div].r52 += cost;

  // 3. Zonas
  const isRck = zona.includes('RCK') || zona.includes('RACK');
  const isRhb = zona.includes('RHB') || zona.includes('HIGHBAY') || (!isRck && zona.length > 0);
  const targetZ = isRck ? acc.zoneAgg.rck : (isRhb ? acc.zoneAgg.rhb : null);

  if (targetZ) {
    targetZ.totalCost += cost;
    targetZ.bultos += bultos;
    targetZ.lpns += lpnVal;
    if (rng === '0 a 10 Semanas') targetZ.r010 += cost;
    else if (rng === '10 a 25 Semanas') targetZ.r1025 += cost;
    else if (rng === '25 a 52 Semanas') targetZ.r2552 += cost;
    else if (rng === 'mayor a 52 Semanas') {
      targetZ.r52 += cost;
      if (sku) {
        if (!targetZ.skus52[sku]) {
          targetZ.skus52[sku] = { sku, desc, fv: rngFv, lpns: 0, cost: 0, bultos: 0 };
        }
        targetZ.skus52[sku].cost += cost;
        targetZ.skus52[sku].bultos += bultos;
        targetZ.skus52[sku].lpns += lpnVal;
      }
    }
  }

  // 4. Top SKUs >52 Semanas
  if (rng === 'mayor a 52 Semanas' && sku) {
    if (!acc.skusOver52[sku]) {
      acc.skusOver52[sku] = { sku, desc, div, rngFv, lpns: 0, cost: 0, bultos: 0, onHand: 0, diasTotal: 0, diasCount: 0 };
    }
    acc.skusOver52[sku].cost += cost;
    acc.skusOver52[sku].bultos += bultos;
    acc.skusOver52[sku].onHand += onHand;
    acc.skusOver52[sku].lpns += lpnVal;
    acc.skusOver52[sku].diasTotal += dias;
    acc.skusOver52[sku].diasCount++;
  }

  // 5. Ubicaciones
  if (ubic) {
    if (!acc.locationsAgg[div]) {
      acc.locationsAgg[div] = { div, r010Ubic: new Set(), r1025Ubic: new Set(), r2552Ubic: new Set(), r52Ubic: new Set(), totalUbic: new Set() };
    }
    acc.locationsAgg[div].totalUbic.add(ubic);
    if (rng === '0 a 10 Semanas') acc.locationsAgg[div].r010Ubic.add(ubic);
    else if (rng === '10 a 25 Semanas') acc.locationsAgg[div].r1025Ubic.add(ubic);
    else if (rng === '25 a 52 Semanas') acc.locationsAgg[div].r2552Ubic.add(ubic);
    else if (rng === 'mayor a 52 Semanas') acc.locationsAgg[div].r52Ubic.add(ubic);
  }
}

async function compileDataOptimized(dataRows, colMap, evolSecos, evolFrescos, activeWeekTag) {
  const CHUNK_SIZE = 8000;
  const totalRows = dataRows.length;

  const acc655 = initAccumulator('655', 'CD Secos 655');
  const acc676 = initAccumulator('676', 'CD Frescos 676');

  for (let i = 0; i < totalRows; i += CHUNK_SIZE) {
    const end = Math.min(i + CHUNK_SIZE, totalRows);

    for (let r = i; r < end; r++) {
      const row = dataRows[r];
      if (!row || row.length === 0) continue;

      const rawWhse = colMap.whse !== -1 ? String(row[colMap.whse] || '').trim() : '';
      let targetAcc = null;
      if (rawWhse === '655' || rawWhse.toLowerCase().includes('seco')) targetAcc = acc655;
      else if (rawWhse === '676' || rawWhse.toLowerCase().includes('fresco')) targetAcc = acc676;
      else continue;

      accumulateRow(row, colMap, targetAcc);
    }

    const pct = Math.round((end / totalRows) * 100);
    showLoading(`Procesando filas por lotes optimizados...`, `Fila ${end.toLocaleString()} de ${totalRows.toLocaleString()} (${pct}%)`);
    await yieldToUi();
  }

  const secos = finalizeWarehouseData(acc655, evolSecos, activeWeekTag);
  const frescos = finalizeWarehouseData(acc676, evolFrescos, activeWeekTag);

  return { secos, frescos };
}

function finalizeWarehouseData(acc, evolObj, activeWeekTag) {
  const { whseCode, whseLabel, totalCost, totalBultos, totalLpns, rangeAgg, divisionAgg, zoneAgg, skusOver52, locationsAgg } = acc;

  if (totalCost === 0 && totalBultos === 0) {
    return createEmptyWarehouseData(whseCode, whseLabel, evolObj, activeWeekTag);
  }

  // Estructuración de Rangos con % BULTOS, % COSTO y % LPN con precisión matemática exacta (5 decimales)
  const ranges = Object.keys(rangeAgg).map(k => {
    const itm = rangeAgg[k];
    const bultosPct = totalBultos > 0 ? Number(((itm.bultos / totalBultos) * 100).toFixed(5)) : 0;
    const costPct = totalCost > 0 ? Number(((itm.cost / totalCost) * 100).toFixed(5)) : 0;
    const lpnsPct = totalLpns > 0 ? Number(((itm.lpns / totalLpns) * 100).toFixed(5)) : 0;

    return {
      label: k,
      cost: itm.cost,
      costPct: costPct,
      lpns: itm.lpns,
      lpnsPct: lpnsPct,
      bultos: itm.bultos,
      bultosPct: bultosPct,
      color: itm.color,
      status: itm.status
    };
  });

  // Divisiones
  const divisions = Object.values(divisionAgg)
    .sort((a, b) => b.total - a.total)
    .map(d => ({
      code: d.code,
      r010: d.r010,
      r1025: d.r1025,
      r2552: d.r2552,
      r52: d.r52,
      total: d.total,
      pct: totalCost > 0 ? parseFloat(((d.total / totalCost) * 100).toFixed(1)) : 0
    }));

  // Top 10 SKUs >52s con Suma de LPNs
  const top10Skus = Object.values(skusOver52)
    .sort((a, b) => b.cost - a.cost)
    .slice(0, 10)
    .map((s, idx) => ({
      sku: s.sku,
      desc: s.desc,
      div: s.div,
      rngFv: s.rngFv,
      lpns: s.lpns,
      cost: Math.round(s.cost),
      bultos: s.bultos,
      onHand: s.onHand,
      days: s.diasCount > 0 ? Math.round(s.diasTotal / s.diasCount) : 0,
      badge: idx === 0 ? 'CRÍTICO #1' : (idx === 1 ? 'CRÍTICO #2' : (idx < 5 ? 'PRIORIDAD' : 'SALDO'))
    }));

  // Zonas RCK vs RHB con Suma de LPNs
  const zones = {
    rck: {
      totalCost: zoneAgg.rck.totalCost,
      lpns: zoneAgg.rck.lpns,
      bultos: zoneAgg.rck.bultos,
      pct: totalCost > 0 ? parseFloat(((zoneAgg.rck.totalCost / totalCost) * 100).toFixed(1)) : 0,
      r010: zoneAgg.rck.r010,
      r1025: zoneAgg.rck.r1025,
      r2552: zoneAgg.rck.r2552,
      r52: zoneAgg.rck.r52,
      lpns52: Object.values(zoneAgg.rck.skus52).reduce((sum, item) => sum + item.lpns, 0),
      topSkus: Object.values(zoneAgg.rck.skus52)
        .sort((a, b) => b.cost - a.cost)
        .slice(0, 10)
        .map(s => ({ sku: s.sku, desc: s.desc, fv: s.fv, lpns: s.lpns, cost: Math.round(s.cost), bultos: s.bultos }))
    },
    rhb: {
      totalCost: zoneAgg.rhb.totalCost,
      lpns: zoneAgg.rhb.lpns,
      bultos: zoneAgg.rhb.bultos,
      pct: totalCost > 0 ? parseFloat(((zoneAgg.rhb.totalCost / totalCost) * 100).toFixed(1)) : 0,
      r010: zoneAgg.rhb.r010,
      r1025: zoneAgg.rhb.r1025,
      r2552: zoneAgg.rhb.r2552,
      r52: zoneAgg.rhb.r52,
      lpns52: Object.values(zoneAgg.rhb.skus52).reduce((sum, item) => sum + item.lpns, 0),
      topSkus: Object.values(zoneAgg.rhb.skus52)
        .sort((a, b) => b.cost - a.cost)
        .slice(0, 10)
        .map(s => ({ sku: s.sku, desc: s.desc, fv: s.fv, lpns: s.lpns, cost: Math.round(s.cost), bultos: s.bultos }))
    }
  };

  // Ubicaciones
  let locationsTotal = 0;
  Object.values(locationsAgg).forEach(l => locationsTotal += l.totalUbic.size);

  const locations = Object.values(locationsAgg)
    .sort((a, b) => b.totalUbic.size - a.totalUbic.size)
    .map(l => ({
      div: l.div,
      r010: l.r010Ubic.size,
      r1025: l.r1025Ubic.size,
      r2552: l.r2552Ubic.size,
      r52: l.r52Ubic.size,
      total: l.totalUbic.size,
      pct: locationsTotal > 0 ? parseFloat(((l.totalUbic.size / locationsTotal) * 100).toFixed(1)) : 0
    }));

  // Sincronización de 7 Semanas del Evolutivo
  const currentWeekBultosPcts = {
    '0 a 10 Semanas': ranges[0].bultosPct,
    '10 a 25 Semanas': ranges[1].bultosPct,
    '25 a 52 Semanas': ranges[2].bultosPct,
    'mayor a 52 Semanas': ranges[3].bultosPct
  };

  let finalWeeks = [];
  let finalEvolution = [];
  const defaultRangeColors = {
    '0 a 10 Semanas': '#10b981',
    '10 a 25 Semanas': '#f59e0b',
    '25 a 52 Semanas': '#f97316',
    'Mayor a 52 Semanas': '#ef4444'
  };

  if (evolObj && evolObj.weeks && evolObj.weeks.length > 0) {
    const evolWeeks = [...evolObj.weeks];
    const lastWeekInSheet = evolWeeks[evolWeeks.length - 1];

    if (lastWeekInSheet === activeWeekTag) {
      finalWeeks = evolWeeks.slice(-7);
      finalEvolution = evolObj.evolution.map(e => {
        const vals = [...(e.values || [])].slice(-7);
        const normLabel = normalizeRange(e.label);
        if (currentWeekBultosPcts[normLabel] !== undefined && vals.length > 0) {
          vals[vals.length - 1] = currentWeekBultosPcts[normLabel];
        }
        return {
          label: e.label,
          values: vals,
          color: e.color || defaultRangeColors[e.label] || '#2563eb'
        };
      });
    } else {
      const hist6Weeks = evolWeeks.slice(-6);
      finalWeeks = [...hist6Weeks, activeWeekTag];

      finalEvolution = evolObj.evolution.map(e => {
        const hist6Vals = (e.values || []).slice(-6);
        const normLabel = normalizeRange(e.label);
        const currentPct = currentWeekBultosPcts[normLabel] !== undefined ? currentWeekBultosPcts[normLabel] : 0;
        return {
          label: e.label,
          values: [...hist6Vals, currentPct],
          color: e.color || defaultRangeColors[e.label] || '#2563eb'
        };
      });
    }
  } else {
    finalWeeks = ['S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', activeWeekTag];
    finalEvolution = [
      { label: '0 a 10 Semanas', values: [89.14, 87.07, 85.83, 86.07, 87.09, 87.96, currentWeekBultosPcts['0 a 10 Semanas']], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [8.73, 10.68, 11.83, 11.71, 10.74, 10.18, currentWeekBultosPcts['10 a 25 Semanas']], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [1.83, 1.95, 2.03, 1.94, 1.91, 1.60, currentWeekBultosPcts['25 a 52 Semanas']], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.30, 0.31, 0.31, 0.28, 0.26, 0.27, currentWeekBultosPcts['mayor a 52 Semanas']], color: '#ef4444' }
    ];
  }

  return {
    name: whseLabel,
    whseCode,
    badgeText: whseCode === '655' ? 'CD SECOS 655' : 'CD FRESCOS 676',
    totalCost: Math.round(totalCost),
    totalLpns,
    totalBultos: Math.round(totalBultos),
    ranges,
    divisions,
    top10Skus,
    zones,
    weeks: finalWeeks,
    evolution: finalEvolution,
    locationsTotal,
    locations
  };
}

function createEmptyWarehouseData(whseTarget, whseLabel, evolObj, activeWeekTag = 'S-37') {
  const weeks = evolObj ? evolObj.weeks : ['S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', activeWeekTag];
  const evolution = evolObj ? evolObj.evolution : [
    { label: '0 a 10 Semanas', values: [0, 0, 0, 0, 0, 0, 0], color: '#10b981' },
    { label: '10 a 25 Semanas', values: [0, 0, 0, 0, 0, 0, 0], color: '#f59e0b' },
    { label: '25 a 52 Semanas', values: [0, 0, 0, 0, 0, 0, 0], color: '#f97316' },
    { label: 'Mayor a 52 Semanas', values: [0, 0, 0, 0, 0, 0, 0], color: '#ef4444' }
  ];

  return {
    name: whseLabel,
    whseCode: whseTarget,
    badgeText: whseTarget === '655' ? 'CD SECOS 655' : 'CD FRESCOS 676',
    totalCost: 0,
    totalLpns: 0,
    totalBultos: 0,
    ranges: [
      { label: '0 a 10 Semanas', cost: 0, costPct: 0, lpns: 0, lpnsPct: 0, bultos: 0, bultosPct: 0, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 0, costPct: 0, lpns: 0, lpnsPct: 0, bultos: 0, bultosPct: 0, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 0, costPct: 0, lpns: 0, lpnsPct: 0, bultos: 0, bultosPct: 0, color: '#f97316', status: 'Riesgo Medio' },
      { label: 'mayor a 52 Semanas', cost: 0, costPct: 0, lpns: 0, lpnsPct: 0, bultos: 0, bultosPct: 0, color: '#ef4444', status: 'Crítico >1 año' }
    ],
    divisions: [],
    top10Skus: [],
    zones: {
      rck: { totalCost: 0, lpns: 0, bultos: 0, pct: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, lpns52: 0, topSkus: [] },
      rhb: { totalCost: 0, lpns: 0, bultos: 0, pct: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, lpns52: 0, topSkus: [] }
    },
    weeks,
    evolution,
    locationsTotal: 0,
    locations: []
  };
}

// ══════════════════════════════════════════════════════════════════════════════
// 7. DATOS DEMOSTRATIVOS (SEMANA 37)
// ══════════════════════════════════════════════════════════════════════════════

function loadSampleData() {
  showLoading('Cargando datos demostrativos de la Semana 37...');
  currentWeekLabel = 'Semana 37';

  // CD Secos 655
  ACTIVE_DATABASE['655'] = {
    name: 'CD Huachipa - Secos',
    whseCode: '655',
    badgeText: 'CD SECOS 655',
    totalCost: 112812540,
    totalLpns: 33265,
    totalBultos: 1345343,
    ranges: [
      { label: '0 a 10 Semanas', cost: 97924406, costPct: 86.80, lpns: 27817, lpnsPct: 83.62, bultos: 1183316, bultosPct: 87.96, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 12502561, costPct: 11.08, lpns: 4471, lpnsPct: 13.44, bultos: 136958, bultosPct: 10.18, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 2096734, costPct: 1.86, lpns: 899, lpnsPct: 2.70, bultos: 21460, bultosPct: 1.60, color: '#f97316', status: 'Riesgo Medio' },
      { label: 'mayor a 52 Semanas', cost: 288839, costPct: 0.26, lpns: 78, lpnsPct: 0.23, bultos: 3609, bultosPct: 0.27, color: '#ef4444', status: 'Crítico >1 año' }
    ],
    divisions: [
      { code: 'J01-PGC COMESTIBLE', r010: 44893121, r1025: 2409955, r2552: 414515, r52: 167255, total: 47884847, pct: 42.4 },
      { code: 'J08-VESTUARIO', r010: 12452170, r1025: 3546454, r2552: 220136, r52: 0, total: 16218760, pct: 14.4 },
      { code: 'J09-HOGAR', r010: 11679599, r1025: 2343584, r2552: 576331, r52: 81400, total: 14680914, pct: 13.0 },
      { code: 'J02-PGC NO COMESTIBLE', r010: 11873814, r1025: 2012474, r2552: 236322, r52: 7779, total: 14130389, pct: 12.5 },
      { code: 'J11-ELECTROHOGAR', r010: 11857252, r1025: 911068, r2552: 126627, r52: 0, total: 12894946, pct: 11.4 },
      { code: 'J10-BAZAR', r010: 2213932, r1025: 1180948, r2552: 486869, r52: 7664, total: 3889414, pct: 3.4 },
      { code: 'J06-PANADERIA Y PASTELERIA', r010: 1074602, r1025: 90837, r2552: 21380, r52: 1405, total: 1188224, pct: 1.1 },
      { code: 'J12-INSTITUCIONALES', r010: 887262, r1025: 0, r2552: 4370, r52: 872, total: 892505, pct: 0.8 },
      { code: 'J05-FLC', r010: 879846, r1025: 2994, r2552: 0, r52: 0, total: 882840, pct: 0.8 },
      { code: 'J07-PLATOS PREPARADOS', r010: 112807, r1025: 4247, r2552: 10184, r52: 22463, total: 149701, pct: 0.1 }
    ],
    top10Skus: [
      { sku: '43314915', desc: 'SET X 2 ESPECIERO TAPA CORCHO 90ML', div: 'J09-HOGAR', rngFv: '-', lpns: 5, cost: 10227, bultos: 138, onHand: 3312, days: 492, badge: 'SALDO' },
      { sku: '43491111', desc: 'COMBO MUG APILABLE VERANO CJ', div: 'J09-HOGAR', rngFv: '-', lpns: 3, cost: 10663, bultos: 85, onHand: 85, days: 569, badge: 'SALDO' },
      { sku: '41843146', desc: 'KETCHUP AMERICANO TOTTUS X 425GR', div: 'J01-PGC COMESTIBLE', rngFv: '19/05/2027', lpns: 3, cost: 12653, bultos: 198, onHand: 3168, days: 473, badge: 'SALDO' },
      { sku: '43488563', desc: 'COMBO GUANTE CON SILIC Y TELA 2025', div: 'J09-HOGAR', rngFv: '-', lpns: 5, cost: 17027, bultos: 117, onHand: 117, days: 512, badge: 'PRIORIDAD' },
      { sku: '42464523', desc: 'PULPA FINAL DE TOMATE TOTTUS X 400 G', div: 'J01-PGC COMESTIBLE', rngFv: '30/09/2027', lpns: 6, cost: 26206, bultos: 785, onHand: 9420, days: 368, badge: 'PRIORIDAD' },
      { sku: '43111173', desc: 'DURAZNO EN MITADES PRECIO UNO 415G', div: 'J01-PGC COMESTIBLE', rngFv: '20/07/2027', lpns: 20, cost: 61831, bultos: 1110, onHand: 26640, days: 682, badge: 'CRÍTICO #1' },
      { sku: '43491112', desc: 'COMBO MUG APILABLE VERANO PU', div: 'J09-HOGAR', rngFv: '-', lpns: 7, cost: 26720, bultos: 213, onHand: 213, days: 569, badge: 'PRIORIDAD' },
      { sku: '43439251', desc: 'LAMINA DE AJI S IMPRES 280MM PET PE 60U', div: 'J07-PLATOS PREPARADOS', rngFv: '05/08/2026', lpns: 1, cost: 11671, bultos: 55, onHand: 55, days: 395, badge: 'SALDO' },
      { sku: '43438965', desc: 'VINO TINTO ALBACORA X750ML', div: 'J01-PGC COMESTIBLE', rngFv: '10/10/2028', lpns: 4, cost: 54671, bultos: 323, onHand: 1938, days: 550, badge: 'CRÍTICO #2' },
      { sku: '43439250', desc: 'LAMINA DE AJI S IMPRES 170MM PET PE 60U', div: 'J07-PLATOS PREPARADOS', rngFv: '05/08/2026', lpns: 1, cost: 10792, bultos: 73, onHand: 73, days: 395, badge: 'SALDO' }
    ],
    zones: {
      rck: {
        totalCost: 34441661,
        lpns: 7219,
        bultos: 360574,
        pct: 30.5,
        r010: 29270343,
        r1025: 4468255,
        r2552: 522709,
        r52: 180353,
        lpns52: 40,
        topSkus: [
          { sku: '43111173', desc: 'DURAZNO EN MITADES PRECIO UNO 415G', fv: '20/07/2027', lpns: 20, cost: 61831, bultos: 1110 },
          { sku: '43438965', desc: 'VINO TINTO ALBACORA X750ML', fv: '10/10/2028', lpns: 4, cost: 54671, bultos: 323 },
          { sku: '42464523', desc: 'PULPA FINAL DE TOMATE TOTTUS X 400 G', fv: '30/09/2027', lpns: 6, cost: 26206, bultos: 785 },
          { sku: '43439251', desc: 'LAMINA DE AJI S IMPRES 280MM PET PE 60U', fv: '05/08/2026', lpns: 1, cost: 11671, bultos: 55 },
          { sku: '43439250', desc: 'LAMINA DE AJI S IMPRES 170MM PET PE 60U', fv: '05/08/2026', lpns: 1, cost: 10792, bultos: 73 }
        ]
      },
      rhb: {
        totalCost: 78367955,
        lpns: 26043,
        bultos: 984765,
        pct: 69.5,
        r010: 68654062,
        r1025: 8034306,
        r2552: 1571100,
        r52: 108486,
        lpns52: 38,
        topSkus: [
          { sku: '43491112', desc: 'COMBO MUG APILABLE VERANO PU', fv: '-', lpns: 7, cost: 26720, bultos: 213 },
          { sku: '43488563', desc: 'COMBO GUANTE CON SILIC Y TELA 2025', fv: '-', lpns: 5, cost: 17027, bultos: 117 },
          { sku: '41843146', desc: 'KETCHUP AMERICANO TOTTUS X 425GR', fv: '19/05/2027', lpns: 3, cost: 12653, bultos: 198 },
          { sku: '43491111', desc: 'COMBO MUG APILABLE VERANO CJ', fv: '-', lpns: 3, cost: 10663, bultos: 85 },
          { sku: '43314915', desc: 'SET X 2 ESPECIERO TAPA CORCHO 90ML', fv: '-', lpns: 5, cost: 10227, bultos: 138 }
        ]
      }
    },
    weeks: ['S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', 'S-37'],
    evolution: [
      { label: '0 a 10 Semanas', values: [89.14, 87.07, 85.83, 86.07, 87.09, 87.96, 87.96], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [8.73, 10.68, 11.83, 11.71, 10.74, 10.18, 10.18], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [1.83, 1.95, 2.03, 1.94, 1.91, 1.60, 1.60], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.30, 0.31, 0.31, 0.28, 0.26, 0.27, 0.27], color: '#ef4444' }
    ],
    locationsTotal: 22872,
    locations: [
      { div: 'J01-PGC COMESTIBLE', r010: 7892, r1025: 660, r2552: 68, r52: 35, total: 8274, pct: 36.2 },
      { div: 'J09-HOGAR', r010: 5863, r1025: 1604, r2552: 477, r52: 31, total: 7435, pct: 32.5 },
      { div: 'J08-VESTUARIO', r010: 3030, r1025: 597, r2552: 34, r52: 0, total: 3632, pct: 15.9 },
      { div: 'J02-PGC NO COMESTIBLE', r010: 2257, r1025: 658, r2552: 67, r52: 2, total: 2724, pct: 11.9 },
      { div: 'J10-BAZAR', r010: 1266, r1025: 605, r2552: 188, r52: 6, total: 2025, pct: 8.9 },
      { div: 'J11-ELECTROHOGAR', r010: 1178, r1025: 147, r2552: 30, r52: 0, total: 1346, pct: 5.9 },
      { div: 'Otras (J05, J06, J07, J12)', r010: 702, r1025: 50, r2552: 18, r52: 4, total: 974, pct: 4.3 }
    ]
  };

  // CD Frescos 676
  ACTIVE_DATABASE['676'] = {
    name: 'CD Villa El Salvador - Frescos',
    whseCode: '676',
    badgeText: 'CD FRESCOS 676',
    totalCost: 7869125,
    totalLpns: 2514,
    totalBultos: 122522,
    ranges: [
      { label: '0 a 10 Semanas', cost: 7467994, costPct: 94.90, lpns: 2393, lpnsPct: 95.19, bultos: 117009, bultosPct: 95.50, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 334413, costPct: 4.25, lpns: 87, lpnsPct: 3.46, bultos: 4662, bultosPct: 3.81, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 56368, costPct: 0.72, lpns: 24, lpnsPct: 0.95, bultos: 642, bultosPct: 0.52, color: '#f97316', status: 'Riesgo Medio' },
      { label: 'mayor a 52 Semanas', cost: 10350, costPct: 0.13, lpns: 10, lpnsPct: 0.40, bultos: 209, bultosPct: 0.17, color: '#ef4444', status: 'Crítico >1 año' }
    ],
    divisions: [
      { code: 'J03-CARNES Y PESCADOS', r010: 2085627, r1025: 90246, r2552: 6893, r52: 584, total: 2183351, pct: 27.7 },
      { code: 'J05-FLC', r010: 1493357, r1025: 82019, r2552: 9193, r52: 0, total: 1584568, pct: 20.1 },
      { code: 'J06-PANADERIA Y PASTELERIA', r010: 1026325, r1025: 16732, r2552: 3887, r52: 2134, total: 1049078, pct: 13.3 },
      { code: 'J01-PGC COMESTIBLE', r010: 953238, r1025: 69648, r2552: 0, r52: 0, total: 1022886, pct: 13.0 },
      { code: 'J04-FRUTAS Y VERDURAS', r010: 1015684, r1025: 0, r2552: 0, r52: 0, total: 1015684, pct: 12.9 },
      { code: 'J07-PLATOS PREPARADOS', r010: 893762, r1025: 75769, r2552: 36395, r52: 7632, total: 1013559, pct: 12.9 }
    ],
    top10Skus: [
      { sku: '43829328', desc: 'ENVASE BISAGRA 121 TR PET', div: 'J07-PLATOS PREPARADOS', rngFv: '-', lpns: 4, cost: 5194, bultos: 53, onHand: 10600, days: 1203, badge: 'CRÍTICO #1' },
      { sku: '42583778', desc: 'ESTUCHE MULTIUSO 46 H60 D100 TR PET', div: 'J07-PLATOS PREPARADOS', rngFv: '-', lpns: 3, cost: 2438, bultos: 16, onHand: 3840, days: 1564, badge: 'CRÍTICO #2' },
      { sku: '42794726', desc: 'ESTUCHE MULTIUSO 46 H60 D100 TR PET', div: 'J06-PANADERIA Y PASTELERIA', rngFv: '-', lpns: 2, cost: 2134, bultos: 14, onHand: 3360, days: 1440, badge: 'PRIORIDAD' },
      { sku: '48866028', desc: 'PIERNITAS DE POLLO IMP TT X KG', div: 'J03-CARNES Y PESCADOS', rngFv: '13/03/2027 - 18/06/2028', lpns: 3, cost: 584, bultos: 126, onHand: 126, days: 488, badge: 'PRIORIDAD' }
    ],
    zones: {
      rck: {
        totalCost: 3147650,
        lpns: 1005,
        bultos: 49008,
        pct: 40.0,
        r010: 2987197,
        r1025: 133765,
        r2552: 22547,
        r52: 4140,
        lpns52: 4,
        topSkus: [
          { sku: '43829328', desc: 'ENVASE BISAGRA 121 TR PET', fv: '-', lpns: 4, cost: 5194, bultos: 53 }
        ]
      },
      rhb: {
        totalCost: 4721475,
        lpns: 1509,
        bultos: 73514,
        pct: 60.0,
        r010: 4480797,
        r1025: 200648,
        r2552: 33821,
        r52: 6210,
        lpns52: 6,
        topSkus: [
          { sku: '42583778', desc: 'ESTUCHE MULTIUSO 46 H60 D100 TR PET', fv: '-', lpns: 3, cost: 2438, bultos: 16 },
          { sku: '42794726', desc: 'ESTUCHE MULTIUSO 46 H60 D100 TR PET', fv: '-', lpns: 2, cost: 2134, bultos: 14 }
        ]
      }
    },
    weeks: ['S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', 'S-37'],
    evolution: [
      { label: '0 a 10 Semanas', values: [95.20, 94.80, 94.10, 93.90, 94.20, 94.50, 95.50], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [4.20, 4.50, 5.10, 5.30, 5.00, 4.80, 3.81], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [0.50, 0.60, 0.70, 0.70, 0.70, 0.60, 0.52], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.17], color: '#ef4444' }
    ],
    locationsTotal: 2514,
    locations: [
      { div: 'J03-CARNES Y PESCADOS', r010: 672, r1025: 28, r2552: 2, r52: 1, total: 703, pct: 28.0 },
      { div: 'J05-FLC', r010: 504, r1025: 25, r2552: 3, r52: 0, total: 532, pct: 21.2 },
      { div: 'J06-PANADERIA Y PASTELERIA', r010: 340, r1025: 6, r2552: 2, r52: 1, total: 349, pct: 13.9 },
      { div: 'J01-PGC COMESTIBLE', r010: 320, r1025: 22, r2552: 0, r52: 0, total: 342, pct: 13.6 },
      { div: 'J04-FRUTAS Y VERDURAS', r010: 325, r1025: 0, r2552: 0, r52: 0, total: 325, pct: 12.9 },
      { div: 'J07-PLATOS PREPARADOS', r010: 232, r1025: 6, r2552: 17, r52: 8, total: 263, pct: 10.5 }
    ]
  };

  document.getElementById('dropzoneBox').style.display = 'none';
  document.getElementById('fileStatusBar').style.display = 'flex';
  document.getElementById('loadedFileName').textContent = 'Datos Demostrativos (Semana 37)';
  document.getElementById('loadedFileMeta').textContent = 'Valores calculados de ejemplo para CD Secos 655 y Frescos 676';

  updateHeaderWeekBadges();
  renderWarehouseData(currentWarehouse);
  hideLoading();
  showToast('Datos de ejemplo cargados con éxito', 'success', 3000);
}

// ══════════════════════════════════════════════════════════════════════════════
// 8. RENDERIZADO VISUAL DEL DASHBOARD
// ══════════════════════════════════════════════════════════════════════════════

function updateHeaderWeekBadges() {
  ['slide1-week-badge', 'slide2-week-badge', 'slide3-week-badge'].forEach(id => {
    const badge = document.getElementById(id);
    if (badge) badge.innerHTML = `<i class="fa-regular fa-calendar"></i> ${currentWeekLabel}`;
  });
}

function switchWarehouse(whseCode) {
  currentWarehouse = whseCode;
  document.getElementById('btnWhseSecos').classList.toggle('active', whseCode === '655');
  document.getElementById('btnWhseFrescos').classList.toggle('active', whseCode === '676');
  renderWarehouseData(whseCode);
  showToast(`Mostrando ${whseCode === '655' ? 'CD Secos (655)' : 'CD Frescos (676)'}`, 'info');
}

function switchSlide(slideNum) {
  currentSlide = slideNum;
  [1, 2, 3].forEach(n => {
    document.getElementById(`tabBtnSlide${n}`).classList.toggle('active', n === slideNum);
    const container = document.getElementById(`slide${n}-container`);
    if (container) container.style.display = (n === slideNum) ? 'block' : 'none';
  });

  setTimeout(() => {
    if (slideNum === 1 && chartInstanceS1) chartInstanceS1.resize();
    if (slideNum === 3 && chartInstanceS3) chartInstanceS3.resize();
  }, 60);
}

function renderWarehouseData(whseCode) {
  const data = ACTIVE_DATABASE[whseCode];
  if (!data) return;

  ['slide1', 'slide2', 'slide3'].forEach(id => {
    const el = document.getElementById(`${id}-whse-badge`);
    if (el) el.textContent = data.badgeText;
  });

  renderSlide1(data);
  renderSlide2(data);
  renderSlide3(data);
}

function renderSlide1(data) {
  document.getElementById('s1-kpi-total').textContent = formatCurrency(data.totalCost);
  document.getElementById('s1-kpi-lpns').textContent = `${formatNumber(data.totalLpns)} LPNs`;
  document.getElementById('s1-kpi-bultos').textContent = `${formatNumber(data.totalBultos)} Bultos`;

  const r010 = (data.ranges && data.ranges[0]) ? data.ranges[0] : { costPct: 0, cost: 0, bultos: 0, bultosPct: 0 };
  const r1025 = (data.ranges && data.ranges[1]) ? data.ranges[1] : { costPct: 0, cost: 0, bultos: 0, bultosPct: 0, lpns: 0 };
  const r2552 = (data.ranges && data.ranges[2]) ? data.ranges[2] : { cost: 0, costPct: 0, bultos: 0, bultosPct: 0 };
  const r52 = (data.ranges && data.ranges[3]) ? data.ranges[3] : { cost: 0, costPct: 0, bultos: 0, bultosPct: 0, lpns: 0 };

  const critCost = (r2552.cost || 0) + (r52.cost || 0);
  const critPct = (r2552.costPct || 0) + (r52.costPct || 0);
  const critBultos = (r2552.bultos || 0) + (r52.bultos || 0);
  const critBultosPct = (r2552.bultosPct || 0) + (r52.bultosPct || 0);

  // Tarjeta 2: Saludable (0 a 10 Sem) - % Bultos destacado + % Costo
  document.getElementById('s1-kpi-healthy-pct').innerHTML = `${safeFixed(r010.bultosPct, 2, '%')} <span style="font-size:0.80rem; font-weight:800; color:#065f46;">Bultos</span>`;
  document.getElementById('s1-kpi-healthy-cost').textContent = `${safeFixed(r010.costPct, 1, '%')} Costo (${formatCurrency(r010.cost)})`;
  document.getElementById('s1-kpi-healthy-bultos').textContent = `${formatNumber(r010.bultos)} Bultos`;

  // Tarjeta 3: En Seguimiento (10 a 25 Sem)
  document.getElementById('s1-kpi-warn-pct').innerHTML = `${safeFixed(r1025.bultosPct, 2, '%')} <span style="font-size:0.80rem; font-weight:800; color:#92400e;">Bultos</span>`;
  document.getElementById('s1-kpi-warn-cost').textContent = `${safeFixed(r1025.costPct, 1, '%')} Costo (${formatCurrency(r1025.cost)})`;
  document.getElementById('s1-kpi-warn-lpns').textContent = `${formatNumber(r1025.bultos)} Bultos`;

  // Tarjeta 4: Crítico (>25 Sem)
  document.getElementById('s1-kpi-crit-cost').innerHTML = `${safeFixed(critBultosPct, 2, '%')} <span style="font-size:0.80rem; font-weight:800; color:#991b1b;">Bultos</span>`;
  document.getElementById('s1-kpi-crit-pct').textContent = `${safeFixed(critPct, 1, '%')} Costo (${formatCurrency(critCost)})`;
  document.getElementById('s1-kpi-over52').textContent = `>52s: ${formatCurrency(r52.cost)} (${formatNumber(r52.bultos)} Bultos)`;

  // 1. Gráfico S1 con valores encima de las barras
  renderChartS1(data.ranges);

  // 2. Tabla 1 S1: Resumen de Rangos (Table 1 de la presentación original)
  renderTableS1Ranges(data.ranges, data.totalCost, data.totalLpns, data.totalBultos);

  // 3. Tabla 2 S1: Matriz Divisiones (Detalle exacto de PPT y Tabla Dinámica)
  const tbodyDiv = document.getElementById('tbodyS1Divisions');
  tbodyDiv.innerHTML = '';
  if (data.divisions && data.divisions.length > 0) {
    data.divisions.forEach(d => {
      const p010 = r010.cost > 0 ? (d.r010 / r010.cost) * 100 : 0;
      const p1025 = r1025.cost > 0 ? (d.r1025 / r1025.cost) * 100 : 0;
      const p2552 = r2552.cost > 0 ? (d.r2552 / r2552.cost) * 100 : 0;
      const p52 = r52.cost > 0 ? (d.r52 / r52.cost) * 100 : 0;
      const pTot = data.totalCost > 0 ? (d.total / data.totalCost) * 100 : 0;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-weight:700; color:#0f172a; text-align:left;">${d.code}</td>
        <td>${d.r010 > 0 ? formatNumber(d.r010) : '-'}</td>
        <td style="color:#059669; font-weight:600;">${d.r010 > 0 ? (p010 < 1 ? safeFixed(p010, 1, '%') : safeFixed(p010, 0, '%')) : '0%'}</td>
        <td>${d.r1025 > 0 ? formatNumber(d.r1025) : '-'}</td>
        <td style="color:#d97706; font-weight:600;">${d.r1025 > 0 ? (p1025 < 1 ? safeFixed(p1025, 1, '%') : safeFixed(p1025, 0, '%')) : '0%'}</td>
        <td style="${d.r2552 > 100000 ? 'background:#fffbeb; font-weight:700; color:#b45309;' : ''}">${d.r2552 > 0 ? formatNumber(d.r2552) : '-'}</td>
        <td style="color:#ea580c; font-weight:600;">${d.r2552 > 0 ? (p2552 < 1 ? safeFixed(p2552, 1, '%') : safeFixed(p2552, 0, '%')) : '0%'}</td>
        <td style="${d.r52 > 10000 ? 'background:#fee2e2; font-weight:800; color:#dc2626;' : ''}">${d.r52 > 0 ? formatNumber(d.r52) : '-'}</td>
        <td style="color:#dc2626; font-weight:700;">${d.r52 > 0 ? (p52 < 1 ? safeFixed(p52, 1, '%') : safeFixed(p52, 0, '%')) : '0%'}</td>
        <td style="font-weight:800; color:#0f172a;">${formatNumber(d.total)}</td>
        <td style="color:#2563eb; font-weight:800;">${pTot < 1 ? safeFixed(pTot, 2, '%') : safeFixed(pTot, 0, '%')}</td>
      `;
      tbodyDiv.appendChild(tr);
    });
  } else {
    tbodyDiv.innerHTML = `<tr><td colspan="11" style="text-align:center; padding:12px; color:#64748b;">No hay divisiones registradas</td></tr>`;
  }

  document.getElementById('tfootS1Divisions').innerHTML = `
    <tr>
      <td style="text-align:left; font-weight:900;">Total general</td>
      <td style="font-weight:900;">${formatNumber(r010.cost)}</td>
      <td style="font-weight:900; color:#059669;">100%</td>
      <td style="font-weight:900;">${formatNumber(r1025.cost)}</td>
      <td style="font-weight:900; color:#d97706;">100%</td>
      <td style="font-weight:900; color:#ea580c;">${formatNumber(r2552.cost)}</td>
      <td style="font-weight:900; color:#ea580c;">100%</td>
      <td style="font-weight:900; color:#dc2626;">${formatNumber(r52.cost)}</td>
      <td style="font-weight:900; color:#dc2626;">100%</td>
      <td style="color:#0f172a; font-weight:900;">${formatNumber(data.totalCost)}</td>
      <td style="color:#2563eb; font-weight:900;">100%</td>
    </tr>
  `;

  // 4. Tabla 3 S1: Top 10 SKUs >52 Semanas (Suma de LPNs)
  const tbodyTop = document.getElementById('tbodyS1Top10');
  tbodyTop.innerHTML = '';
  let sumTopCost = 0, sumTopLpns = 0, sumTopBultos = 0, sumTopOnHand = 0;

  if (!data.top10Skus || data.top10Skus.length === 0) {
    tbodyTop.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:16px; color:#64748b; font-weight:600;">No se registraron SKUs con más de 52 semanas de antigüedad</td></tr>`;
  } else {
    data.top10Skus.forEach((sku, idx) => {
      sumTopCost += (sku.cost || 0);
      sumTopLpns += (sku.lpns || 0);
      sumTopBultos += (sku.bultos || 0);
      sumTopOnHand += (sku.onHand || 0);

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family:'JetBrains Mono',monospace; font-weight:800; color:#1e293b;">${sku.sku}</td>
        <td style="text-align:left; font-weight:700; color:#0f172a;">${sku.desc}</td>
        <td style="text-align:left; color:#475569; font-weight:600;">${sku.div}</td>
        <td style="text-align:center; color:#64748b; font-size:0.75rem;">${sku.rngFv}</td>
        <td style="font-weight:800; color:#2563eb;">${formatNumber(sku.lpns)}</td>
        <td style="font-weight:800; color:${sku.cost > 20000 ? '#dc2626' : '#0f172a'};">${formatNumber(sku.cost)}</td>
        <td style="font-weight:600;">${formatNumber(sku.bultos)}</td>
        <td style="color:#64748b;">${formatNumber(sku.onHand)}</td>
        <td style="font-weight:700; color:${sku.days > 500 ? '#b91c1c' : '#475569'};">${sku.days} d</td>
        <td style="text-align:center;">
          <span style="background:${idx < 2 ? '#fee2e2' : '#f1f5f9'}; color:${idx < 2 ? '#b91c1c' : '#475569'}; padding:2px 6px; border-radius:5px; font-weight:800; font-size:0.72rem;">
            ${sku.badge}
          </span>
        </td>
      `;
      tbodyTop.appendChild(tr);
    });
  }

  document.getElementById('tfootS1Top10').innerHTML = `
    <tr>
      <td colspan="4" style="text-align:left; font-weight:900;">Total Top 10 SKUs</td>
      <td style="color:#2563eb; font-weight:900;">${formatNumber(sumTopLpns)}</td>
      <td style="color:#dc2626; font-size:0.90rem; font-weight:900;">S/ ${formatNumber(sumTopCost)}</td>
      <td style="font-weight:800;">${formatNumber(sumTopBultos)}</td>
      <td style="font-weight:700; color:#64748b;">${formatNumber(sumTopOnHand)}</td>
      <td colspan="2" style="text-align:center; color:#64748b; font-weight:700;">Concentración crítica</td>
    </tr>
  `;
}

function renderTableS1Ranges(ranges, totalCost, totalLpns, totalBultos) {
  const tbody = document.getElementById('tbodyS1Ranges');
  const tfoot = document.getElementById('tfootS1Ranges');
  if (!tbody || !tfoot) return;

  tbody.innerHTML = '';
  (ranges || []).forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align:left; font-weight:700; color:#0f172a;">
        <span style="display:inline-block; width:8px; height:8px; border-radius:50%; background:${r.color}; margin-right:6px;"></span>
        ${r.label}
      </td>
      <td style="font-weight:700; color:#2563eb;">${formatNumber(r.lpns)}</td>
      <td style="font-weight:600; color:#475569;">${safeFixed(r.lpnsPct, 2, '%')}</td>
      <td>${formatNumber(r.bultos)}</td>
      <td style="font-weight:800; color:${r.color};">${safeFixed(r.bultosPct, 2, '%')}</td>
      <td style="font-weight:700;">${formatCurrency(r.cost)}</td>
      <td style="font-weight:800;">${safeFixed(r.costPct, 1, '%')}</td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td style="text-align:left; font-weight:900;">Total general</td>
      <td style="color:#2563eb; font-weight:900;">${formatNumber(totalLpns)}</td>
      <td style="font-weight:900; color:#475569;">100.00%</td>
      <td style="font-weight:900;">${formatNumber(totalBultos)}</td>
      <td style="font-weight:900;">100%</td>
      <td style="color:#0f172a; font-weight:900;">${formatCurrency(totalCost)}</td>
      <td style="font-weight:900;">100%</td>
    </tr>
  `;
}

// ══════════════════════════════════════════════════════════════════════════════
// 9. GRÁFICOS CHART.JS CON DATALABELS VISIBLES ENCIMA DE LAS BARRAS
// ══════════════════════════════════════════════════════════════════════════════

function setChartS1Mode(mode) {
  currentChartS1Mode = mode;
  const btnBultos = document.getElementById('btnChartModeBultos');
  const btnCosto = document.getElementById('btnChartModeCosto');
  const title = document.getElementById('chartS1Title');

  if (btnBultos && btnCosto) {
    if (mode === 'bultos') {
      btnBultos.style.background = '#ffffff';
      btnBultos.style.color = '#1e3a8a';
      btnBultos.style.fontWeight = '800';
      btnBultos.style.boxShadow = '0 1px 2px rgba(0,0,0,0.08)';

      btnCosto.style.background = 'transparent';
      btnCosto.style.color = '#64748b';
      btnCosto.style.fontWeight = '700';
      btnCosto.style.boxShadow = 'none';

      if (title) title.textContent = 'Distribución de Bultos por Antigüedad';
    } else {
      btnCosto.style.background = '#ffffff';
      btnCosto.style.color = '#1e3a8a';
      btnCosto.style.fontWeight = '800';
      btnCosto.style.boxShadow = '0 1px 2px rgba(0,0,0,0.08)';

      btnBultos.style.background = 'transparent';
      btnBultos.style.color = '#64748b';
      btnBultos.style.fontWeight = '700';
      btnBultos.style.boxShadow = 'none';

      if (title) title.textContent = 'Distribución de Costo por Antigüedad';
    }
  }

  const data = ACTIVE_DATABASE[currentWarehouse];
  if (data && data.ranges) {
    renderChartS1(data.ranges);
  }
}
window.setChartS1Mode = setChartS1Mode;

function renderChartS1(ranges) {
  const ctx = document.getElementById('chartS1Ranges');
  if (!ctx) return;

  if (chartInstanceS1) {
    chartInstanceS1.destroy();
  }

  const isBultos = (currentChartS1Mode === 'bultos');
  const labels = ranges.map(r => r.label);
  const colors = ranges.map(r => r.color);

  // Datos según modo activo:
  const chartData = isBultos
    ? ranges.map(r => r.bultos)
    : ranges.map(r => {
        return r.cost >= 1000000 ? parseFloat((r.cost / 1000000).toFixed(2)) : parseFloat((r.cost / 1000000).toFixed(3));
      });

  // Plugin personalizado para pintar valores y porcentajes ENCIMA de cada barra
  const valueLabelsPlugin = {
    id: 'barValueLabelsOnTop',
    afterDatasetsDraw(chart) {
      const { ctx: c } = chart;
      c.save();
      const meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data) return;

      meta.data.forEach((bar, idx) => {
        const rItem = ranges[idx];
        if (!bar || !rItem) return;

        c.textAlign = 'center';
        c.textBaseline = 'bottom';

        if (isBultos) {
          // Línea 1: Bultos (ej. 117,009 Bultos)
          const bultosText = `${formatNumber(rItem.bultos)} Bultos`;
          c.font = 'bold 11px Inter, sans-serif';
          c.fillStyle = '#0f172a';
          c.fillText(bultosText, bar.x, bar.y - 14);

          // Línea 2: Porcentaje de Bultos (ej. 95.50%)
          const pctText = safeFixed(rItem.bultosPct, 2, '%');
          c.font = '800 11px Inter, sans-serif';
          c.fillStyle = rItem.color || '#2563eb';
          c.fillText(`(${pctText})`, bar.x, bar.y - 2);
        } else {
          // Línea 1: Costo (S/ 7.5M o S/ 334K)
          const costText = formatCompact(rItem.cost);
          c.font = 'bold 12px Inter, sans-serif';
          c.fillStyle = '#0f172a';
          c.fillText(`S/ ${costText}`, bar.x, bar.y - 14);

          // Línea 2: Porcentaje de Costo (ej. 94.9%)
          const pctText = safeFixed(rItem.costPct, 1, '%');
          c.font = '800 11px Inter, sans-serif';
          c.fillStyle = rItem.color || '#2563eb';
          c.fillText(`(${pctText})`, bar.x, bar.y - 2);
        }
      });
      c.restore();
    }
  };

  chartInstanceS1 = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: isBultos ? 'Bultos' : 'Costo (Millones S/)',
        data: chartData,
        backgroundColor: colors,
        borderRadius: 8,
        borderWidth: 0
      }]
    },
    plugins: [valueLabelsPlugin],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: {
          top: 32 // Margen superior para que el texto encima de la barra más alta nunca se corte
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (item) => {
              const r = ranges[item.dataIndex];
              if (!r) return '';
              if (isBultos) {
                return ` ${formatNumber(r.bultos)} Bultos (${safeFixed(r.bultosPct, 2, '%')})`;
              } else {
                return ` ${formatCurrency(r.cost)} (${safeFixed(r.costPct, 1, '%')})`;
              }
            }
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: '30%', // 30% de altura adicional sobre la barra mayor para alojar los datalabels
          grid: { color: '#f1f5f9' },
          ticks: {
            callback: (v) => isBultos ? formatCompact(v) : `S/ ${v}M`,
            font: { size: 11, weight: '700' }
          }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 11, weight: '700' } }
        }
      }
    }
  });
}

function renderSlide2(data) {
  const z = data.zones || { rck: {}, rhb: {} };

  // RCK
  document.getElementById('s2-rck-cost').textContent = formatCurrency(z.rck?.totalCost);
  document.getElementById('s2-rck-sub').textContent = `${safeFixed(z.rck?.pct, 1, '% del CD')} | ${formatNumber(z.rck?.lpns)} LPNs`;
  document.getElementById('s2-rck-010').textContent = formatCompact(z.rck?.r010);
  document.getElementById('s2-rck-1025').textContent = formatCompact(z.rck?.r1025);
  document.getElementById('s2-rck-2552').textContent = formatCompact(z.rck?.r2552);
  document.getElementById('s2-rck-52').textContent = formatNumber(z.rck?.r52);

  // RHB
  document.getElementById('s2-rhb-cost').textContent = formatCurrency(z.rhb?.totalCost);
  document.getElementById('s2-rhb-sub').textContent = `${safeFixed(z.rhb?.pct, 1, '% del CD')} | ${formatNumber(z.rhb?.lpns)} LPNs`;
  document.getElementById('s2-rhb-010').textContent = formatCompact(z.rhb?.r010);
  document.getElementById('s2-rhb-1025').textContent = formatCompact(z.rhb?.r1025);
  document.getElementById('s2-rhb-2552').textContent = formatCompact(z.rhb?.r2552);
  document.getElementById('s2-rhb-52').textContent = formatNumber(z.rhb?.r52);

  // RCK Tablas (Suma de LPNs)
  const tbodyRck = document.getElementById('tbodyS2Rck');
  tbodyRck.innerHTML = '';
  let sumRckCost = 0, sumRckLpns = 0, sumRckBultos = 0;

  if (!z.rck?.topSkus || z.rck.topSkus.length === 0) {
    tbodyRck.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:12px; color:#64748b; font-weight:600;">No hay SKUs >52s en RCK</td></tr>`;
  } else {
    z.rck.topSkus.forEach(s => {
      sumRckCost += (s.cost || 0);
      sumRckLpns += (s.lpns || 0);
      sumRckBultos += (s.bultos || 0);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family:'JetBrains Mono',monospace; font-weight:800;">${s.sku}</td>
        <td style="text-align:left; font-weight:700; color:#0f172a;">${s.desc}</td>
        <td style="text-align:center; color:#64748b; font-size:0.75rem;">${s.fv}</td>
        <td style="font-weight:800; color:#dc2626;">${formatNumber(s.lpns)}</td>
        <td style="font-weight:800; color:#dc2626;">${formatNumber(s.cost)}</td>
        <td style="font-weight:600;">${formatNumber(s.bultos)}</td>
      `;
      tbodyRck.appendChild(tr);
    });
  }

  document.getElementById('tfootS2Rck').innerHTML = `
    <tr>
      <td colspan="3" style="text-align:left; font-weight:900;">Total RCK Top SKUs</td>
      <td style="color:#dc2626; font-weight:900;">${formatNumber(sumRckLpns)}</td>
      <td style="color:#dc2626; font-size:0.90rem; font-weight:900;">S/ ${formatNumber(sumRckCost)}</td>
      <td style="font-weight:800;">${formatNumber(sumRckBultos)}</td>
    </tr>
  `;

  // RHB Tablas (Suma de LPNs)
  const tbodyRhb = document.getElementById('tbodyS2Rhb');
  tbodyRhb.innerHTML = '';
  let sumRhbCost = 0, sumRhbLpns = 0, sumRhbBultos = 0;

  if (!z.rhb?.topSkus || z.rhb.topSkus.length === 0) {
    tbodyRhb.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:12px; color:#64748b; font-weight:600;">No hay SKUs >52s en RHB</td></tr>`;
  } else {
    z.rhb.topSkus.forEach(s => {
      sumRhbCost += (s.cost || 0);
      sumRhbLpns += (s.lpns || 0);
      sumRhbBultos += (s.bultos || 0);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family:'JetBrains Mono',monospace; font-weight:800;">${s.sku}</td>
        <td style="text-align:left; font-weight:700; color:#0f172a;">${s.desc}</td>
        <td style="text-align:center; color:#64748b; font-size:0.75rem;">${s.fv}</td>
        <td style="font-weight:800; color:#2563eb;">${formatNumber(s.lpns)}</td>
        <td style="font-weight:800; color:#1e293b;">${formatNumber(s.cost)}</td>
        <td style="font-weight:600;">${formatNumber(s.bultos)}</td>
      `;
      tbodyRhb.appendChild(tr);
    });
  }

  document.getElementById('tfootS2Rhb').innerHTML = `
    <tr>
      <td colspan="3" style="text-align:left; font-weight:900;">Total RHB Top SKUs</td>
      <td style="color:#2563eb; font-weight:900;">${formatNumber(sumRhbLpns)}</td>
      <td style="color:#0f172a; font-size:0.90rem; font-weight:900;">S/ ${formatNumber(sumRhbCost)}</td>
      <td style="font-weight:800;">${formatNumber(sumRhbBultos)}</td>
    </tr>
  `;
}

function renderSlide3(data) {
  renderChartS3(data.weeks, data.evolution);

  const theadRow = document.getElementById('theadS3EvolRow');
  theadRow.innerHTML = `<th style="text-align:left; font-weight:900;">Rango Semanas</th>`;
  (data.weeks || []).forEach((w, idx) => {
    const isCurrent = idx === data.weeks.length - 1;
    theadRow.innerHTML += `<th style="${isCurrent ? 'background:#eff6ff; color:#1d4ed8; font-weight:900; font-size:0.90rem;' : 'font-weight:800;'}">${w}</th>`;
  });
  theadRow.innerHTML += `<th style="font-weight:900;">Var. WoW</th>`;

  const tbodyEvol = document.getElementById('tbodyS3Evolution');
  tbodyEvol.innerHTML = '';
  (data.evolution || []).forEach(e => {
    const vals = e.values || [];
    const lastVal = vals[vals.length - 1] || 0;
    const prevVal = vals[vals.length - 2] || lastVal;
    const diff = lastVal - prevVal;
    const isNeutral = Math.abs(diff) < 0.005;
    const isPositiveGood = e.label.includes('0 a 10') ? diff > 0 : diff < 0;

    let trHtml = `
      <td style="text-align:left; font-weight:800; color:#1e293b; font-size:0.90rem;">
        <span style="display:inline-block; width:8px; height:8px; border-radius:50%; background:${e.color}; margin-right:6px;"></span>
        ${e.label}
      </td>
    `;
    vals.forEach((v, idx) => {
      const isCurrent = idx === vals.length - 1;
      trHtml += `<td style="${isCurrent ? 'background:#eff6ff; font-weight:900; font-size:0.96rem; color:' + e.color + ';' : 'font-size:0.88rem; font-weight:600;'}">${safeFixed(v, 2, '%')}</td>`;
    });

    trHtml += `
      <td style="font-weight:800; font-size:0.88rem; color:${isNeutral ? '#64748b' : (isPositiveGood ? '#059669' : '#dc2626')};">
        ${isNeutral ? '= 0.00%' : (diff > 0 ? '▲ +' : '▼ -') + safeFixed(Math.abs(diff), 2, '%')}
      </td>
    `;

    const tr = document.createElement('tr');
    tr.innerHTML = trHtml;
    tbodyEvol.appendChild(tr);
  });

  // Ubicaciones
  document.getElementById('s3-total-ubic').textContent = `${formatNumber(data.locationsTotal)} ubicaciones`;
  const tbodyLoc = document.getElementById('tbodyS3Locations');
  tbodyLoc.innerHTML = '';
  let sum010 = 0, sum1025 = 0, sum2552 = 0, sum52 = 0;

  (data.locations || []).forEach(loc => {
    sum010 += (loc.r010 || 0);
    sum1025 += (loc.r1025 || 0);
    sum2552 += (loc.r2552 || 0);
    sum52 += (loc.r52 || 0);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight:700; color:#0f172a;">${loc.div}</td>
      <td style="font-weight:600;">${formatNumber(loc.r010)}</td>
      <td style="${loc.r1025 > 500 ? 'background:#fffbeb; font-weight:700; color:#b45309;' : 'font-weight:600;'}">${formatNumber(loc.r1025)}</td>
      <td style="${loc.r2552 > 100 ? 'background:#ffedd5; font-weight:800; color:#c2410c;' : 'font-weight:600;'}">${formatNumber(loc.r2552)}</td>
      <td style="${loc.r52 > 10 ? 'background:#fee2e2; font-weight:900; color:#dc2626;' : 'font-weight:600;'}">${loc.r52 > 0 ? loc.r52 : '-'}</td>
      <td style="font-weight:800; color:#0f172a;">${formatNumber(loc.total)}</td>
      <td style="color:#2563eb; font-weight:800;">${safeFixed(loc.pct, 1, '%')}</td>
    `;
    tbodyLoc.appendChild(tr);
  });

  document.getElementById('tfootS3Locations').innerHTML = `
    <tr>
      <td style="text-align:left; font-weight:900;">Total Ubicaciones</td>
      <td style="font-weight:900;">${formatNumber(sum010)}</td>
      <td style="font-weight:900;">${formatNumber(sum1025)}</td>
      <td style="color:#ea580c; font-weight:900;">${formatNumber(sum2552)}</td>
      <td style="color:#dc2626; font-weight:900;">${formatNumber(sum52)}</td>
      <td style="color:#0f172a; font-weight:900;">${formatNumber(data.locationsTotal)}</td>
      <td style="color:#2563eb; font-weight:900;">100%</td>
    </tr>
  `;
}

function renderChartS3(weeks, evolution) {
  const ctx = document.getElementById('chartS3Evolution');
  if (!ctx) return;

  if (chartInstanceS3) {
    chartInstanceS3.destroy();
  }

  const wList = weeks || [];
  const datasets = (evolution || []).map(e => ({
    label: e.label,
    data: e.values || [],
    borderColor: e.color,
    backgroundColor: e.color,
    borderWidth: e.label.includes('0 a 10') ? 3 : 2,
    tension: 0.25,
    pointRadius: 4,
    pointHoverRadius: 6
  }));

  chartInstanceS3 = new Chart(ctx, {
    type: 'line',
    data: {
      labels: wList,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'top',
          labels: { boxWidth: 12, font: { size: 11, weight: '700' } }
        },
        tooltip: {
          callbacks: {
            label: (item) => ` ${item.dataset.label}: ${safeFixed(item.raw, 2, '%')} (${safeFixed(item.raw, 4, '%')})`
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          ticks: {
            callback: (v) => `${v}%`,
            font: { size: 11, weight: '700' }
          },
          grid: { color: '#f1f5f9' }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 11, weight: '700' } }
        }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 10. MOTOR DE COPIA 16:9 NATIVO DIRECTO AL PORTAPAPELES (SIN DESCARGAS FORZADAS)
// ══════════════════════════════════════════════════════════════════════════════

async function copyCurrentSlideToClipboard() {
  if (isCapturing) return;
  isCapturing = true;

  const btn = document.getElementById('btnCapturePpt');
  const originalText = btn ? btn.innerHTML : '';
  if (btn) btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Copiando...';

  try {
    const containerId = `slide${currentSlide}-container`;
    const container = document.getElementById(containerId);
    if (!container) throw new Error('Contenedor de diapositiva no encontrado');

    showToast(`📸 Generando diapositiva 16:9 de Lámina ${currentSlide}...`, 'info', 2000);

    // 1. Aplicar clase que bloquea dimensiones estrictas 16:9 (1600x900)
    container.classList.add('capturing-16-9');
    if (currentSlide === 1 && chartInstanceS1) chartInstanceS1.resize();
    if (currentSlide === 3 && chartInstanceS3) chartInstanceS3.resize();
    await yieldToUi();

    // 2. Renderizar con html2canvas en alta definición
    const renderedCanvas = await window.html2canvas(container, {
      scale: 1.5, // Resolución nítida Retina
      backgroundColor: '#ffffff',
      useCORS: true,
      logging: false
    });

    // Restaurar vista web responsiva inmediatamente
    container.classList.remove('capturing-16-9');
    if (currentSlide === 1 && chartInstanceS1) chartInstanceS1.resize();
    if (currentSlide === 3 && chartInstanceS3) chartInstanceS3.resize();

    // 3. Crear canvas panorámico exacto 1920x1080 (16:9)
    const W = 1920;
    const H = 1080;
    const offscreen = document.createElement('canvas');
    offscreen.width = W;
    offscreen.height = H;
    const ctx = offscreen.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);

    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(renderedCanvas, 0, 0, W, H);

    lastCapturedDataUrl = offscreen.toDataURL('image/png');

    // 4. Escribir directamente al Portapapeles (Clipboard API)
    offscreen.toBlob(async (blob) => {
      if (!blob) {
        showToast('Error al crear imagen de la diapositiva', 'danger');
        return;
      }

      lastCapturedBlob = blob;

      if (navigator.clipboard && window.ClipboardItem) {
        try {
          const item = new ClipboardItem({ 'image/png': blob });
          await navigator.clipboard.write([item]);
          showToast(`✅ ¡Lámina ${currentSlide} copiada! Lista para pegar en PowerPoint (Ctrl + V)`, 'success', 5000);
          showCopySuccessModal(lastCapturedDataUrl);
        } catch (clipErr) {
          console.warn('Acceso directo al portapapeles restringido:', clipErr);
          // Si el navegador bloquea la escritura automática, mostrar modal con botón explícito
          showCopySuccessModal(lastCapturedDataUrl);
        }
      } else {
        showCopySuccessModal(lastCapturedDataUrl);
      }
    }, 'image/png');

  } catch (err) {
    console.error('Error al capturar diapositiva:', err);
    showToast('No se pudo generar la diapositiva', 'danger');
  } finally {
    if (btn) btn.innerHTML = originalText;
    setTimeout(() => { isCapturing = false; }, 300);
  }
}

function showCopySuccessModal(dataUrl) {
  const modal = document.getElementById('modalCopySuccess');
  const img = document.getElementById('modalPreviewImg');
  if (img) img.src = dataUrl;
  if (modal) modal.style.display = 'flex';
}

function closeCopyModal() {
  const modal = document.getElementById('modalCopySuccess');
  if (modal) modal.style.display = 'none';
}

async function copyModalImgAgain() {
  if (!lastCapturedBlob) return;
  try {
    const item = new ClipboardItem({ 'image/png': lastCapturedBlob });
    await navigator.clipboard.write([item]);
    showToast('✅ ¡Copiado nuevamente al portapapeles!', 'success', 3000);
  } catch (err) {
    showToast('Por favor pulsa botón derecho sobre la imagen y selecciona "Copiar imagen"', 'info', 4500);
  }
}

function downloadModalImg() {
  if (!lastCapturedDataUrl) return;
  const link = document.createElement('a');
  link.download = `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`;
  link.href = lastCapturedDataUrl;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

async function downloadCurrentSlide() {
  const containerId = `slide${currentSlide}-container`;
  const container = document.getElementById(containerId);
  if (!container) return;

  showToast(`💾 Generando archivo PNG 16:9 (1920x1080)...`, 'info', 2000);

  container.classList.add('capturing-16-9');
  await yieldToUi();

  const renderedCanvas = await window.html2canvas(container, {
    scale: 1.5,
    backgroundColor: '#ffffff'
  });

  container.classList.remove('capturing-16-9');

  const W = 1920;
  const H = 1080;
  const offscreen = document.createElement('canvas');
  offscreen.width = W;
  offscreen.height = H;
  const ctx = offscreen.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, W, H);
  ctx.drawImage(renderedCanvas, 0, 0, W, H);

  const link = document.createElement('a');
  link.download = `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`;
  link.href = offscreen.toDataURL('image/png');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
