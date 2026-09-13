/**
 * DataNexus - Antigüedad de Inventario (Secos 655 y Frescos 676)
 * Motor Ultra-Robusto de Procesamiento de Excel (.xlsx, .xlsm, .xlsb)
 * Resuelve automáticamente encabezados con espacios, diferentes filas de inicio y formatos numéricos.
 */

// Estado global
let currentWarehouse = '655'; // '655' (Secos) o '676' (Frescos)
let currentSlide = 1;
let currentWeekLabel = 'Semana Actual';
let chartInstanceS1 = null;
let chartInstanceS3 = null;
let isCapturing = false;

// Almacén de datos activos
let ACTIVE_DATABASE = {
  '655': null,
  '676': null
};

// ══════════════════════════════════════════════════════════════════════════════
// 1. UTILIDADES Y PROTECCIÓN CONTRA ERRORES DE FORMATO
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Función segura para aplicar toFixed sin riesgo de TypeError
 */
function safeFixed(val, decimals = 1, suffix = '') {
  if (val === null || val === undefined || isNaN(Number(val))) {
    return '0.0' + suffix;
  }
  return Number(val).toFixed(decimals) + suffix;
}

function parseNum(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  if (!v) return 0;
  const clean = String(v).replace(/,/g, '').replace(/[^\d\.-]/g, '').trim();
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
  if (n >= 1000000) return (n / 1000000).toFixed(2) + 'M';
  if (n >= 1000) return (n / 1000).toFixed(1) + 'K';
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

function showLoading(text) {
  const overlay = document.getElementById('loadingOverlay');
  const txt = document.getElementById('loadingText');
  if (txt) txt.textContent = text || 'Procesando archivo Excel...';
  if (overlay) overlay.style.display = 'flex';
}

function hideLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) overlay.style.display = 'none';
}

// ══════════════════════════════════════════════════════════════════════════════
// 2. BUSCADOR INTELIGENTE DE COLUMNAS (NORMALIZADOR FUZZY)
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Busca un valor en un objeto de fila sin importar mayúsculas, minúsculas,
 * tildes o espacios extras en el encabezado de Excel (ej: "WHSE ", " Semana ").
 */
function getRowField(row, candidateNames) {
  if (!row || typeof row !== 'object') return '';
  const rowKeys = Object.keys(row);

  for (const cand of candidateNames) {
    const cleanCand = cand.toLowerCase().replace(/[\s_\-\.\:\/]/g, '').trim();
    for (const key of rowKeys) {
      const cleanKey = key.toLowerCase().replace(/[\s_\-\.\:\/]/g, '').trim();
      if (cleanKey === cleanCand) {
        return row[key];
      }
    }
  }
  return '';
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
// 4. MOTOR PRINCIPAL DE LECTURA DE EXCEL (SHEETJS)
// ══════════════════════════════════════════════════════════════════════════════

function processExcelFile(file) {
  showLoading(`Abriendo archivo: ${file.name}...`);

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      showLoading('Inspeccionando hojas del libro de trabajo...');

      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true,
        cellNF: false,
        cellText: false
      });

      // 1. Detectar Hoja Principal BD
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

      showLoading(`Procesando hoja principal "${sheetNameBD}"...`);

      // Detectar automáticamente la fila exacta donde empiezan los encabezados
      const headerRowIndex = detectHeaderRowIndex(wsBD);
      const rawRows = XLSX.utils.sheet_to_json(wsBD, {
        range: headerRowIndex,
        defval: ''
      });

      if (!rawRows || rawRows.length === 0) {
        throw new Error(`La hoja "${sheetNameBD}" no contiene registros de inventario.`);
      }

      // 2. Detectar Hoja EVOLUTIVO
      let evolSecos = null;
      let evolFrescos = null;

      const sheetNameEvol = workbook.SheetNames.find(n => {
        const clean = n.trim().toUpperCase();
        return clean.includes('EVOL') || clean.includes('HIST');
      });

      if (sheetNameEvol && workbook.Sheets[sheetNameEvol]) {
        showLoading('Extrayendo evolución de semanas desde la hoja EVOLUTIVO...');
        const wsEvol = workbook.Sheets[sheetNameEvol];
        const evolFound = parseEvolutivoSheetSmart(wsEvol);
        evolSecos = evolFound.secos;
        evolFrescos = evolFound.frescos;
      }

      // 3. Compilar datos para CD Secos (655) y CD Frescos (676)
      showLoading('Calculando indicadores de inventario para Secos (655) y Frescos (676)...');
      ACTIVE_DATABASE['655'] = compileWarehouseData(rawRows, '655', 'CD Secos 655', evolSecos);
      ACTIVE_DATABASE['676'] = compileWarehouseData(rawRows, '676', 'CD Frescos 676', evolFrescos);

      // 4. Detectar semana actual de los datos para la cabecera
      const firstRow = rawRows[0] || {};
      const semVal = getRowField(firstRow, ['semana', 'sem', 'week']);
      if (semVal) {
        currentWeekLabel = `Semana ${String(semVal).replace(/[^\d]/g, '') || semVal}`;
      } else if (evolSecos && evolSecos.weeks && evolSecos.weeks.length > 0) {
        currentWeekLabel = `Semana ${evolSecos.weeks[evolSecos.weeks.length - 1].replace(/[^\d]/g, '')}`;
      }

      // 5. Actualizar interfaz
      document.getElementById('dropzoneBox').style.display = 'none';
      document.getElementById('fileStatusBar').style.display = 'flex';
      document.getElementById('loadedFileName').textContent = `Archivo: ${file.name}`;
      document.getElementById('loadedFileMeta').textContent = `${rawRows.length.toLocaleString('en-US')} filas procesadas de "${sheetNameBD}" | ${sheetNameEvol ? `Hoja "${sheetNameEvol}" conectada` : 'Cálculos directos de BD'}`;

      updateHeaderWeekBadges();
      renderWarehouseData(currentWarehouse);
      hideLoading();
      showToast(`¡Archivo ${file.name} procesado con éxito!`, 'success', 4500);

    } catch (err) {
      console.error('Error al procesar Excel:', err);
      hideLoading();
      alert(`Error al leer el archivo Excel: ${err.message}\nPor favor verifica que la hoja "BD" contenga la información de inventario.`);
    }
  };

  reader.onerror = function () {
    hideLoading();
    showToast('Error al leer el archivo desde el disco', 'danger');
  };

  reader.readAsArrayBuffer(file);
}

/**
 * Escanea las primeras 12 filas de la hoja para encontrar la fila que contiene
 * los encabezados clave como "WHSE", "SKU", "COSTOS", "LPN" o "DESCRIPCION".
 */
function detectHeaderRowIndex(ws) {
  if (!ws || !ws['!ref']) return 0;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const maxScanRow = Math.min(range.e.r, 12);

  for (let r = range.s.r; r <= maxScanRow; r++) {
    let rowText = '';
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined) {
        rowText += ' ' + String(cell.v).toUpperCase();
      }
    }
    if ((rowText.includes('WHSE') || rowText.includes('ALMACEN')) &&
        (rowText.includes('SKU') || rowText.includes('COSTO') || rowText.includes('LPN'))) {
      return r;
    }
  }
  return 0;
}

// ══════════════════════════════════════════════════════════════════════════════
// 5. EXTRACCIÓN INTELIGENTE DE LA HOJA EVOLUTIVO
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Escanea la hoja "EVOLUTIVO" buscando las filas de rangos ("0 a 10 Semanas")
 * para Secos y para Frescos, sin depender de que estén rígidamente en la celda C5.
 */
function parseEvolutivoSheetSmart(ws) {
  const result = { secos: null, frescos: null };
  if (!ws || !ws['!ref']) return result;

  const range = XLSX.utils.decode_range(ws['!ref']);
  const matchingRows = [];

  // Buscar filas que contengan "0 a 10"
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

  // Primer bloque encontrado = Secos (655)
  if (matchingRows.length > 0) {
    result.secos = extractEvolutivoBlock(ws, matchingRows[0].rowDataStart, range);
  } else {
    // Fallback a coordenadas tradicionales C5:BC9
    result.secos = extractEvolutivoBlock(ws, 5, range);
  }

  // Segundo bloque encontrado = Frescos (676)
  if (matchingRows.length > 1) {
    result.frescos = extractEvolutivoBlock(ws, matchingRows[1].rowDataStart, range);
  } else {
    // Fallback a coordenadas tradicionales C14:BC18
    result.frescos = extractEvolutivoBlock(ws, 14, range);
  }

  return result;
}

function extractEvolutivoBlock(ws, dataStartRow, range) {
  const headerRow = Math.max(0, dataStartRow - 1);
  const validCols = [];

  // Escaneo horizontal de columnas
  for (let c = 1; c <= range.e.c; c++) {
    const hCell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    const dCell = ws[XLSX.utils.encode_cell({ r: dataStartRow, c })];

    if (hCell && hCell.v !== undefined && String(hCell.v).trim() !== '' &&
        dCell && dCell.v !== undefined && String(dCell.v).trim() !== '') {
      validCols.push(c);
    }
  }

  if (validCols.length === 0) return null;

  // Extraer las últimas 7 semanas con información
  const last7Cols = validCols.slice(-7);

  const weeks = last7Cols.map(c => {
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
    const values = last7Cols.map(c => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v === undefined) return 0;
      let val = Number(cell.v);
      if (isNaN(val)) {
        val = parseFloat(String(cell.v).replace('%', '').replace(',', '.').trim()) || 0;
      } else if (val <= 1.0 && val > 0) {
        val = val * 100; // Formato porcentaje decimal (0.8861 -> 88.61)
      }
      return isNaN(val) ? 0 : parseFloat(val.toFixed(2));
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
// 6. COMPILACIÓN DE DATOS DESDE LAS FILAS DE "tbBD"
// ══════════════════════════════════════════════════════════════════════════════

function compileWarehouseData(rawRows, whseTarget, whseLabel, evolObj) {
  // Filtrar filas por WHSE con tolerancia de espacios y tipos
  const rows = rawRows.filter(r => {
    const wVal = String(getRowField(r, ['whse', 'almacen_fisico', 'almacenfisico', 'almacen'])).trim();
    if (wVal === String(whseTarget)) return true;
    if (whseTarget === '655' && (wVal.toLowerCase().includes('seco') || wVal.includes('655'))) return true;
    if (whseTarget === '676' && (wVal.toLowerCase().includes('fresco') || wVal.includes('676'))) return true;
    return false;
  });

  if (rows.length === 0) {
    return createEmptyWarehouseData(whseTarget, whseLabel, evolObj);
  }

  let totalCost = 0;
  let totalBultos = 0;
  const lpnsSet = new Set();

  const rangeAgg = {
    '0 a 10 Semanas': { cost: 0, bultos: 0, lpns: new Set(), color: '#10b981', status: 'Saludable' },
    '10 a 25 Semanas': { cost: 0, bultos: 0, lpns: new Set(), color: '#f59e0b', status: 'En Alerta' },
    '25 a 52 Semanas': { cost: 0, bultos: 0, lpns: new Set(), color: '#f97316', status: 'Riesgo Medio' },
    'mayor a 52 Semanas': { cost: 0, bultos: 0, lpns: new Set(), color: '#ef4444', status: 'Crítico >1 año' }
  };

  const divisionAgg = {};

  const zoneAgg = {
    rck: { totalCost: 0, lpns: new Set(), bultos: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, skus52: {} },
    rhb: { totalCost: 0, lpns: new Set(), bultos: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, skus52: {} }
  };

  const skusOver52 = {};
  const locationsAgg = {};

  rows.forEach((r, rowIdx) => {
    const cost = parseNum(getRowField(r, ['costos', 'costo', 'costo_total', 'costototal']));
    const bultos = parseNum(getRowField(r, ['bultos', 'bulto', 'cajas', 'cantidad']));
    const onHand = parseNum(getRowField(r, ['on_hand', 'onhand', 'unidades']));
    const dias = parseNum(getRowField(r, ['dias', 'dias_antiguedad', 'antiguedad_dias']));

    const lpn = String(getRowField(r, ['lpn', 'nro_lpn', 'lpns']) || ('LPN_' + rowIdx)).trim();
    const sku = String(getRowField(r, ['sku', 'codigo', 'cod_sku'])).trim();
    const desc = String(getRowField(r, ['descripcion', 'desc', 'descripcion_sku', 'producto'])).trim();
    const div = String(getRowField(r, ['division', 'div', 'dpto', 'departamento']) || 'OTROS').trim();
    const rng = normalizeRange(getRowField(r, ['semanas', 'rango_semanas', 'antiguedad_semanas', 'antiguedad']));
    const zona = String(getRowField(r, ['zona', 'tipo_zona', 'tipozona'])).toUpperCase().trim();
    const ubic = String(getRowField(r, ['ubicacion', 'posicion', 'slot'])).trim();
    const rngFv = String(getRowField(r, ['rng-fv', 'rng_fv', 'rngfv', 'fecha_vencim', 'vencimiento']) || '-').trim();

    totalCost += cost;
    totalBultos += bultos;
    lpnsSet.add(lpn);

    // 1. Acumular Rangos
    if (rangeAgg[rng]) {
      rangeAgg[rng].cost += cost;
      rangeAgg[rng].bultos += bultos;
      rangeAgg[rng].lpns.add(lpn);
    }

    // 2. Acumular Divisiones
    if (!divisionAgg[div]) {
      divisionAgg[div] = { code: div, r010: 0, r1025: 0, r2552: 0, r52: 0, total: 0 };
    }
    divisionAgg[div].total += cost;
    if (rng === '0 a 10 Semanas') divisionAgg[div].r010 += cost;
    else if (rng === '10 a 25 Semanas') divisionAgg[div].r1025 += cost;
    else if (rng === '25 a 52 Semanas') divisionAgg[div].r2552 += cost;
    else if (rng === 'mayor a 52 Semanas') divisionAgg[div].r52 += cost;

    // 3. Acumular Zonas
    const isRck = zona.includes('RCK') || zona.includes('RACK');
    const isRhb = zona.includes('RHB') || zona.includes('HIGHBAY') || (!isRck && zona.length > 0);
    const targetZ = isRck ? zoneAgg.rck : (isRhb ? zoneAgg.rhb : null);

    if (targetZ) {
      targetZ.totalCost += cost;
      targetZ.bultos += bultos;
      targetZ.lpns.add(lpn);
      if (rng === '0 a 10 Semanas') targetZ.r010 += cost;
      else if (rng === '10 a 25 Semanas') targetZ.r1025 += cost;
      else if (rng === '25 a 52 Semanas') targetZ.r2552 += cost;
      else if (rng === 'mayor a 52 Semanas') {
        targetZ.r52 += cost;
        if (sku) {
          if (!targetZ.skus52[sku]) {
            targetZ.skus52[sku] = { sku, desc, fv: rngFv, lpns: new Set(), cost: 0, bultos: 0 };
          }
          targetZ.skus52[sku].cost += cost;
          targetZ.skus52[sku].bultos += bultos;
          targetZ.skus52[sku].lpns.add(lpn);
        }
      }
    }

    // 4. Top SKUs >52 Semanas
    if (rng === 'mayor a 52 Semanas' && sku) {
      if (!skusOver52[sku]) {
        skusOver52[sku] = { sku, desc, div, rngFv, lpns: new Set(), cost: 0, bultos: 0, onHand: 0, diasTotal: 0, diasCount: 0 };
      }
      skusOver52[sku].cost += cost;
      skusOver52[sku].bultos += bultos;
      skusOver52[sku].onHand += onHand;
      skusOver52[sku].lpns.add(lpn);
      skusOver52[sku].diasTotal += dias;
      skusOver52[sku].diasCount++;
    }

    // 5. Ubicaciones
    if (ubic) {
      if (!locationsAgg[div]) {
        locationsAgg[div] = { div, r010Ubic: new Set(), r1025Ubic: new Set(), r2552Ubic: new Set(), r52Ubic: new Set(), totalUbic: new Set() };
      }
      locationsAgg[div].totalUbic.add(ubic);
      if (rng === '0 a 10 Semanas') locationsAgg[div].r010Ubic.add(ubic);
      else if (rng === '10 a 25 Semanas') locationsAgg[div].r1025Ubic.add(ubic);
      else if (rng === '25 a 52 Semanas') locationsAgg[div].r2552Ubic.add(ubic);
      else if (rng === 'mayor a 52 Semanas') locationsAgg[div].r52Ubic.add(ubic);
    }
  });

  // Estructuración de Rangos
  const ranges = Object.keys(rangeAgg).map(k => {
    const itm = rangeAgg[k];
    return {
      label: k,
      cost: itm.cost,
      costPct: totalCost > 0 ? (itm.cost / totalCost) * 100 : 0,
      lpns: itm.lpns.size,
      lpnsPct: lpnsSet.size > 0 ? (itm.lpns.size / lpnsSet.size) * 100 : 0,
      bultos: itm.bultos,
      bultosPct: totalBultos > 0 ? (itm.bultos / totalBultos) * 100 : 0,
      color: itm.color,
      status: itm.status
    };
  });

  // Estructuración de Divisiones
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

  // Top 10 SKUs >52s
  const top10Skus = Object.values(skusOver52)
    .sort((a, b) => b.cost - a.cost)
    .slice(0, 10)
    .map((s, idx) => ({
      sku: s.sku,
      desc: s.desc,
      div: s.div,
      rngFv: s.rngFv,
      lpns: s.lpns.size,
      cost: Math.round(s.cost),
      bultos: s.bultos,
      onHand: s.onHand,
      days: s.diasCount > 0 ? Math.round(s.diasTotal / s.diasCount) : 0,
      badge: idx === 0 ? 'CRÍTICO #1' : (idx === 1 ? 'CRÍTICO #2' : (idx < 5 ? 'PRIORIDAD' : 'SALDO'))
    }));

  // Zonas
  const zones = {
    rck: {
      totalCost: zoneAgg.rck.totalCost,
      lpns: zoneAgg.rck.lpns.size,
      bultos: zoneAgg.rck.bultos,
      pct: totalCost > 0 ? parseFloat(((zoneAgg.rck.totalCost / totalCost) * 100).toFixed(1)) : 0,
      r010: zoneAgg.rck.r010,
      r1025: zoneAgg.rck.r1025,
      r2552: zoneAgg.rck.r2552,
      r52: zoneAgg.rck.r52,
      lpns52: Object.values(zoneAgg.rck.skus52).reduce((sum, item) => sum + item.lpns.size, 0),
      topSkus: Object.values(zoneAgg.rck.skus52)
        .sort((a, b) => b.cost - a.cost)
        .slice(0, 10)
        .map(s => ({ sku: s.sku, desc: s.desc, fv: s.fv, lpns: s.lpns.size, cost: Math.round(s.cost), bultos: s.bultos }))
    },
    rhb: {
      totalCost: zoneAgg.rhb.totalCost,
      lpns: zoneAgg.rhb.lpns.size,
      bultos: zoneAgg.rhb.bultos,
      pct: totalCost > 0 ? parseFloat(((zoneAgg.rhb.totalCost / totalCost) * 100).toFixed(1)) : 0,
      r010: zoneAgg.rhb.r010,
      r1025: zoneAgg.rhb.r1025,
      r2552: zoneAgg.rhb.r2552,
      r52: zoneAgg.rhb.r52,
      lpns52: Object.values(zoneAgg.rhb.skus52).reduce((sum, item) => sum + item.lpns.size, 0),
      topSkus: Object.values(zoneAgg.rhb.skus52)
        .sort((a, b) => b.cost - a.cost)
        .slice(0, 10)
        .map(s => ({ sku: s.sku, desc: s.desc, fv: s.fv, lpns: s.lpns.size, cost: Math.round(s.cost), bultos: s.bultos }))
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

  const weeks = evolObj ? evolObj.weeks : ['S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', 'S-37'];
  const evolution = evolObj ? evolObj.evolution : [
    { label: '0 a 10 Semanas', values: [88.61, 89.14, 87.07, 85.83, 86.07, 87.09, rangeAgg['0 a 10 Semanas'].bultos / (totalBultos || 1) * 100], color: '#10b981' },
    { label: '10 a 25 Semanas', values: [9.47, 8.73, 10.68, 11.83, 11.71, 10.74, rangeAgg['10 a 25 Semanas'].bultos / (totalBultos || 1) * 100], color: '#f59e0b' },
    { label: '25 a 52 Semanas', values: [1.60, 1.83, 1.95, 2.03, 1.94, 1.91, rangeAgg['25 a 52 Semanas'].bultos / (totalBultos || 1) * 100], color: '#f97316' },
    { label: 'Mayor a 52 Semanas', values: [0.32, 0.30, 0.31, 0.31, 0.28, 0.26, rangeAgg['mayor a 52 Semanas'].bultos / (totalBultos || 1) * 100], color: '#ef4444' }
  ];

  return {
    name: whseLabel,
    whseCode: whseTarget,
    badgeText: whseTarget === '655' ? 'CD SECOS 655' : 'CD FRESCOS 676',
    totalCost: Math.round(totalCost),
    totalLpns: lpnsSet.size,
    totalBultos: Math.round(totalBultos),
    ranges,
    divisions,
    top10Skus,
    zones,
    weeks,
    evolution,
    locationsTotal,
    locations
  };
}

function createEmptyWarehouseData(whseTarget, whseLabel, evolObj) {
  const weeks = evolObj ? evolObj.weeks : ['S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', 'S-37'];
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
// 7. RENDERIZADO VISUAL DEL DASHBOARD
// ══════════════════════════════════════════════════════════════════════════════

function updateHeaderWeekBadges() {
  const badge1 = document.getElementById('slide1-week-badge');
  if (badge1) badge1.innerHTML = `<i class="fa-regular fa-calendar"></i> ${currentWeekLabel}`;
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

  const r010 = (data.ranges && data.ranges[0]) ? data.ranges[0] : { costPct: 0, cost: 0, bultosPct: 0 };
  document.getElementById('s1-kpi-healthy-pct').textContent = `${safeFixed(r010.costPct, 1, '%')}`;
  document.getElementById('s1-kpi-healthy-cost').textContent = formatCurrency(r010.cost);
  document.getElementById('s1-kpi-healthy-bultos').textContent = `${safeFixed(r010.bultosPct, 1, '% Bultos')}`;

  const r1025 = (data.ranges && data.ranges[1]) ? data.ranges[1] : { costPct: 0, cost: 0, lpns: 0 };
  document.getElementById('s1-kpi-warn-pct').textContent = `${safeFixed(r1025.costPct, 1, '%')}`;
  document.getElementById('s1-kpi-warn-cost').textContent = formatCurrency(r1025.cost);
  document.getElementById('s1-kpi-warn-lpns').textContent = `${formatNumber(r1025.lpns)} LPNs`;

  const r2552 = (data.ranges && data.ranges[2]) ? data.ranges[2] : { cost: 0, costPct: 0 };
  const r52 = (data.ranges && data.ranges[3]) ? data.ranges[3] : { cost: 0, costPct: 0, lpns: 0 };
  const critCost = (r2552.cost || 0) + (r52.cost || 0);
  const critPct = (r2552.costPct || 0) + (r52.costPct || 0);

  document.getElementById('s1-kpi-crit-cost').textContent = formatCurrency(critCost);
  document.getElementById('s1-kpi-crit-pct').textContent = `${safeFixed(critPct, 1, '% del Capital')}`;
  document.getElementById('s1-kpi-over52').textContent = `>52s: ${formatCompact(r52.cost)} (${formatNumber(r52.lpns)} LPNs)`;

  // Gráfico S1
  renderChartS1(data.ranges);

  // Tabla S1: Matriz Divisiones
  const tbodyDiv = document.getElementById('tbodyS1Divisions');
  tbodyDiv.innerHTML = '';
  if (data.divisions && data.divisions.length > 0) {
    data.divisions.forEach(d => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${d.code}</td>
        <td>${formatCompact(d.r010)}</td>
        <td>${formatCompact(d.r1025)}</td>
        <td style="${d.r2552 > 200000 ? 'background:#fffbeb; font-weight:700; color:#b45309;' : ''}">${formatCompact(d.r2552)}</td>
        <td style="${d.r52 > 10000 ? 'background:#fee2e2; font-weight:800; color:#dc2626;' : ''}">${d.r52 > 0 ? formatNumber(d.r52) : '-'}</td>
        <td style="font-weight:700;">${formatCompact(d.total)}</td>
        <td style="color:#64748b; font-weight:600;">${safeFixed(d.pct, 1, '%')}</td>
      `;
      tbodyDiv.appendChild(tr);
    });
  } else {
    tbodyDiv.innerHTML = `<tr><td colspan="7" style="text-align:center; padding:12px; color:#64748b;">No hay divisiones registradas</td></tr>`;
  }

  document.getElementById('tfootS1Divisions').innerHTML = `
    <tr>
      <td>Total General</td>
      <td>${formatCompact(r010.cost)}</td>
      <td>${formatCompact(r1025.cost)}</td>
      <td style="color:#c2410c;">${formatCompact(r2552.cost)}</td>
      <td style="color:#dc2626;">${formatNumber(r52.cost)}</td>
      <td>${formatCompact(data.totalCost)}</td>
      <td>100%</td>
    </tr>
  `;

  // Tabla S1: Top 10 SKUs >52 Semanas
  const tbodyTop = document.getElementById('tbodyS1Top10');
  tbodyTop.innerHTML = '';
  let sumTopCost = 0, sumTopLpns = 0, sumTopBultos = 0, sumTopOnHand = 0;

  if (!data.top10Skus || data.top10Skus.length === 0) {
    tbodyTop.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:18px; color:#64748b;">No se registraron SKUs con más de 52 semanas de antigüedad</td></tr>`;
  } else {
    data.top10Skus.forEach((sku, idx) => {
      sumTopCost += (sku.cost || 0);
      sumTopLpns += (sku.lpns || 0);
      sumTopBultos += (sku.bultos || 0);
      sumTopOnHand += (sku.onHand || 0);

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family:'JetBrains Mono',monospace; font-weight:700; color:#334155;">${sku.sku}</td>
        <td style="text-align:left; font-weight:600; color:#0f172a;">${sku.desc}</td>
        <td style="text-align:left; color:#64748b;">${sku.div}</td>
        <td style="text-align:center; color:#64748b; font-size:0.7rem;">${sku.rngFv}</td>
        <td style="font-weight:700;">${sku.lpns}</td>
        <td style="font-weight:800; color:${sku.cost > 20000 ? '#dc2626' : '#0f172a'};">${formatNumber(sku.cost)}</td>
        <td>${formatNumber(sku.bultos)}</td>
        <td style="color:#64748b;">${formatNumber(sku.onHand)}</td>
        <td style="font-weight:700; color:${sku.days > 500 ? '#b91c1c' : '#475569'};">${sku.days} d</td>
        <td style="text-align:center;">
          <span style="background:${idx < 2 ? '#fee2e2' : '#f1f5f9'}; color:${idx < 2 ? '#b91c1c' : '#475569'}; padding:2px 6px; border-radius:4px; font-weight:700; font-size:0.68rem;">
            ${sku.badge}
          </span>
        </td>
      `;
      tbodyTop.appendChild(tr);
    });
  }

  document.getElementById('tfootS1Top10').innerHTML = `
    <tr>
      <td colspan="4" style="text-align:left;">Total Top 10 SKUs</td>
      <td>${sumTopLpns}</td>
      <td style="color:#dc2626; font-size:0.82rem;">S/ ${formatNumber(sumTopCost)}</td>
      <td>${formatNumber(sumTopBultos)}</td>
      <td>${formatNumber(sumTopOnHand)}</td>
      <td colspan="2" style="text-align:center; color:#64748b;">Concentración crítica</td>
    </tr>
  `;
}

function renderChartS1(ranges) {
  const ctx = document.getElementById('chartS1Ranges');
  if (!ctx) return;

  if (chartInstanceS1) {
    chartInstanceS1.destroy();
  }

  const rList = ranges || [];
  const labels = rList.map(r => r.label);
  const dataCosts = rList.map(r => safeFixed(r.cost / 1000000, 2));
  const colors = rList.map(r => r.color || '#2563eb');

  chartInstanceS1 = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Costo (Millones S/)',
        data: dataCosts,
        backgroundColor: colors,
        borderRadius: 6,
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (item) => ` S/ ${item.raw} Millones (${safeFixed(rList[item.dataIndex]?.costPct, 1, '%')})`
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grid: { color: '#f1f5f9' },
          ticks: {
            callback: (v) => `S/ ${v}M`,
            font: { size: 10 }
          }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 10, weight: 600 } }
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

  // RCK Tablas
  const tbodyRck = document.getElementById('tbodyS2Rck');
  tbodyRck.innerHTML = '';
  let sumRckCost = 0, sumRckLpns = 0, sumRckBultos = 0;

  if (!z.rck?.topSkus || z.rck.topSkus.length === 0) {
    tbodyRck.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:12px; color:#64748b;">No hay SKUs >52s en RCK</td></tr>`;
  } else {
    z.rck.topSkus.forEach(s => {
      sumRckCost += (s.cost || 0);
      sumRckLpns += (s.lpns || 0);
      sumRckBultos += (s.bultos || 0);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family:'JetBrains Mono',monospace;">${s.sku}</td>
        <td style="text-align:left; font-weight:600;">${s.desc}</td>
        <td style="text-align:center; color:#64748b; font-size:0.7rem;">${s.fv}</td>
        <td style="font-weight:700; color:#dc2626;">${s.lpns}</td>
        <td style="font-weight:800; color:#dc2626;">${formatNumber(s.cost)}</td>
        <td>${formatNumber(s.bultos)}</td>
      `;
      tbodyRck.appendChild(tr);
    });
  }

  document.getElementById('tfootS2Rck').innerHTML = `
    <tr>
      <td colspan="3" style="text-align:left;">Total RCK Top SKUs</td>
      <td style="color:#dc2626;">${sumRckLpns}</td>
      <td style="color:#dc2626; font-size:0.8rem;">S/ ${formatNumber(sumRckCost)}</td>
      <td>${formatNumber(sumRckBultos)}</td>
    </tr>
  `;

  // RHB Tablas
  const tbodyRhb = document.getElementById('tbodyS2Rhb');
  tbodyRhb.innerHTML = '';
  let sumRhbCost = 0, sumRhbLpns = 0, sumRhbBultos = 0;

  if (!z.rhb?.topSkus || z.rhb.topSkus.length === 0) {
    tbodyRhb.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:12px; color:#64748b;">No hay SKUs >52s en RHB</td></tr>`;
  } else {
    z.rhb.topSkus.forEach(s => {
      sumRhbCost += (s.cost || 0);
      sumRhbLpns += (s.lpns || 0);
      sumRhbBultos += (s.bultos || 0);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family:'JetBrains Mono',monospace;">${s.sku}</td>
        <td style="text-align:left; font-weight:600;">${s.desc}</td>
        <td style="text-align:center; color:#64748b; font-size:0.7rem;">${s.fv}</td>
        <td style="font-weight:700; color:#2563eb;">${s.lpns}</td>
        <td style="font-weight:800; color:#1e293b;">${formatNumber(s.cost)}</td>
        <td>${formatNumber(s.bultos)}</td>
      `;
      tbodyRhb.appendChild(tr);
    });
  }

  document.getElementById('tfootS2Rhb').innerHTML = `
    <tr>
      <td colspan="3" style="text-align:left;">Total RHB Top SKUs</td>
      <td style="color:#2563eb;">${sumRhbLpns}</td>
      <td style="color:#0f172a; font-size:0.8rem;">S/ ${formatNumber(sumRhbCost)}</td>
      <td>${formatNumber(sumRhbBultos)}</td>
    </tr>
  `;
}

function renderSlide3(data) {
  renderChartS3(data.weeks, data.evolution);

  const theadRow = document.getElementById('theadS3EvolRow');
  theadRow.innerHTML = `<th style="text-align:left;">Rango Semanas</th>`;
  (data.weeks || []).forEach((w, idx) => {
    const isCurrent = idx === data.weeks.length - 1;
    theadRow.innerHTML += `<th style="${isCurrent ? 'background:#eff6ff; color:#1d4ed8; font-weight:800;' : ''}">${w}</th>`;
  });
  theadRow.innerHTML += `<th>Var. WoW</th>`;

  const tbodyEvol = document.getElementById('tbodyS3Evolution');
  tbodyEvol.innerHTML = '';
  (data.evolution || []).forEach(e => {
    const vals = e.values || [];
    const lastVal = vals[vals.length - 1] || 0;
    const prevVal = vals[vals.length - 2] || lastVal;
    const diff = lastVal - prevVal;
    const isPositiveGood = e.label.includes('0 a 10') ? diff > 0 : diff < 0;

    let trHtml = `
      <td style="text-align:left; font-weight:700; color:#1e293b;">
        <span style="display:inline-block; width:8px; height:8px; border-radius:50%; background:${e.color}; margin-right:6px;"></span>
        ${e.label}
      </td>
    `;
    vals.forEach((v, idx) => {
      const isCurrent = idx === vals.length - 1;
      trHtml += `<td style="${isCurrent ? 'background:#eff6ff; font-weight:800; color:' + e.color + ';' : ''}">${safeFixed(v, 2, '%')}</td>`;
    });

    trHtml += `
      <td style="font-weight:700; color:${isPositiveGood ? '#059669' : '#dc2626'};">
        ${diff >= 0 ? '▲ +' : '▼ '}${safeFixed(Math.abs(diff), 2, '%')}
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
      <td style="font-weight:600;">${loc.div}</td>
      <td>${formatNumber(loc.r010)}</td>
      <td style="${loc.r1025 > 500 ? 'background:#fffbeb; font-weight:700; color:#b45309;' : ''}">${formatNumber(loc.r1025)}</td>
      <td style="${loc.r2552 > 100 ? 'background:#ffedd5; font-weight:700; color:#c2410c;' : ''}">${formatNumber(loc.r2552)}</td>
      <td style="${loc.r52 > 10 ? 'background:#fee2e2; font-weight:800; color:#dc2626;' : ''}">${loc.r52 > 0 ? loc.r52 : '-'}</td>
      <td style="font-weight:700;">${formatNumber(loc.total)}</td>
      <td style="color:#64748b; font-weight:600;">${safeFixed(loc.pct, 1, '%')}</td>
    `;
    tbodyLoc.appendChild(tr);
  });

  document.getElementById('tfootS3Locations').innerHTML = `
    <tr>
      <td style="text-align:left;">Total Ubicaciones</td>
      <td>${formatNumber(sum010)}</td>
      <td>${formatNumber(sum1025)}</td>
      <td style="color:#ea580c;">${formatNumber(sum2552)}</td>
      <td style="color:#dc2626;">${formatNumber(sum52)}</td>
      <td>${formatNumber(data.locationsTotal)}</td>
      <td>100%</td>
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
          labels: { boxWidth: 12, font: { size: 11, weight: 600 } }
        },
        tooltip: {
          callbacks: {
            label: (item) => ` ${item.dataset.label}: ${safeFixed(item.raw, 2, '%')}`
          }
        }
      },
      scales: {
        y: {
          ticks: {
            callback: (v) => `${v}%`,
            font: { size: 10 }
          },
          grid: { color: '#f1f5f9' }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 10, weight: 600 } }
        }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 8. MOTOR DE CAPTURA 16:9 PARA POWERPOINT
// ══════════════════════════════════════════════════════════════════════════════

async function captureCurrentSlide() {
  if (isCapturing) return;

  const containerId = `slide${currentSlide}-container`;
  const container = document.getElementById(containerId);
  if (!container) {
    showToast('No se encontró el contenedor de la diapositiva', 'danger');
    return;
  }

  isCapturing = true;
  showToast(`📸 Generando diapositiva 16:9 de Lámina ${currentSlide}...`, 'info', 2500);

  try {
    const renderedCanvas = await window.html2canvas(container, {
      scale: 2,
      backgroundColor: '#ffffff',
      useCORS: true,
      logging: false
    });

    const W = 1920;
    const H = 1080;
    const offscreen = document.createElement('canvas');
    offscreen.width = W;
    offscreen.height = H;
    const ctx = offscreen.getContext('2d');

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);

    const margin = 40;
    const availW = W - margin * 2;
    const availH = H - margin * 2;

    const scaleFit = Math.min(availW / renderedCanvas.width, availH / renderedCanvas.height);
    const drawW = renderedCanvas.width * scaleFit;
    const drawH = renderedCanvas.height * scaleFit;
    const posX = (W - drawW) / 2;
    const posY = (H - drawH) / 2;

    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(renderedCanvas, posX, posY, drawW, drawH);

    offscreen.toBlob(async (blob) => {
      if (!blob) throw new Error('Error al generar blob de imagen');

      if (navigator.clipboard && window.ClipboardItem) {
        try {
          const item = new ClipboardItem({ 'image/png': blob });
          await navigator.clipboard.write([item]);
          showToast(`✅ ¡Lámina ${currentSlide} copiada en 16:9! Lista para pegar en PowerPoint (Ctrl + V)`, 'success', 5000);
        } catch (clipErr) {
          console.warn('Error al escribir en portapapeles:', clipErr);
          downloadCanvasAsPng(offscreen, `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`);
          showToast(`Descargada como archivo PNG 16:9`, 'info', 4500);
        }
      } else {
        downloadCanvasAsPng(offscreen, `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`);
        showToast(`Descargada como archivo PNG 16:9`, 'info', 4500);
      }
    }, 'image/png');

  } catch (err) {
    console.error('Error al capturar diapositiva:', err);
    showToast('No se pudo generar la captura. Inténtalo nuevamente.', 'danger');
  } finally {
    setTimeout(() => { isCapturing = false; }, 400);
  }
}

async function downloadCurrentSlide() {
  const containerId = `slide${currentSlide}-container`;
  const container = document.getElementById(containerId);
  if (!container) return;

  showToast(`💾 Descargando Lámina ${currentSlide} en 1920x1080...`, 'info', 2000);

  const renderedCanvas = await window.html2canvas(container, {
    scale: 2,
    backgroundColor: '#ffffff'
  });

  const W = 1920;
  const H = 1080;
  const offscreen = document.createElement('canvas');
  offscreen.width = W;
  offscreen.height = H;
  const ctx = offscreen.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, W, H);

  const margin = 40;
  const scaleFit = Math.min((W - margin * 2) / renderedCanvas.width, (H - margin * 2) / renderedCanvas.height);
  const drawW = renderedCanvas.width * scaleFit;
  const drawH = renderedCanvas.height * scaleFit;
  ctx.drawImage(renderedCanvas, (W - drawW) / 2, (H - drawH) / 2, drawW, drawH);

  downloadCanvasAsPng(offscreen, `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`);
}

function downloadCanvasAsPng(canvas, filename) {
  const link = document.createElement('a');
  link.download = filename;
  link.href = canvas.toDataURL('image/png');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
