/**
 * DataNexus - Antigüedad de Inventario (Secos 655 y Frescos 676)
 * Motor de lectura y procesamiento automático de archivos Excel (.xlsx, .xlsm, .xlsb)
 * con renderizado dinámico de KPIs y captura en proporción 16:9 para PowerPoint.
 */

// Estado global
let currentWarehouse = '655'; // '655' (Secos) o '676' (Frescos)
let currentSlide = 1;
let chartInstanceS1 = null;
let chartInstanceS3 = null;
let isCapturing = false;

// Almacén de datos activos (calculados desde Excel o precargados)
let ACTIVE_DATABASE = {
  '655': null,
  '676': null
};

// ══════════════════════════════════════════════════════════════════════════════
// 1. INICIALIZACIÓN Y CONFIGURACIÓN DE DROPZONE
// ══════════════════════════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  setupDropzone();
  // Inicialmente no forzamos datos precargados si no hay archivo, pero dejamos listos los contenedores
});

function setupDropzone() {
  const dropzone = document.getElementById('dropzoneBox');
  if (!dropzone) return;

  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files && files.length > 0) {
      processExcelFile(files[0]);
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
// 2. LECTURA Y PROCESAMIENTO DEL ARCHIVO EXCEL CON SHEETJS
// ══════════════════════════════════════════════════════════════════════════════

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

function processExcelFile(file) {
  showLoading(`Leyendo archivo: ${file.name}...`);

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      showLoading('Analizando hojas "BD" y "EVOLUTIVO"...');

      // Leer libro de Excel
      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true,
        cellNF: false,
        cellText: false
      });

      // 1. Encontrar la hoja "BD" (o que contenga "BD" o tabla "tbBD")
      const sheetNameBD = workbook.SheetNames.find(n => n.trim().toUpperCase() === 'BD') ||
                          workbook.SheetNames.find(n => n.trim().toUpperCase().includes('BD')) ||
                          workbook.SheetNames[0];

      const wsBD = workbook.Sheets[sheetNameBD];
      if (!wsBD) {
        throw new Error('No se encontró la hoja "BD" en el archivo Excel.');
      }

      // Convertir BD a objetos JSON
      showLoading('Extrayendo registros de la base de datos principal...');
      const rawRows = XLSX.utils.sheet_to_json(wsBD, { defval: '' });

      if (!rawRows || rawRows.length === 0) {
        throw new Error('La hoja "BD" está vacía o no contiene filas de datos.');
      }

      // 2. Encontrar y procesar la hoja "EVOLUTIVO"
      const sheetNameEvol = workbook.SheetNames.find(n => n.trim().toUpperCase() === 'EVOLUTIVO');
      let evolSecos = null;
      let evolFrescos = null;

      if (sheetNameEvol && workbook.Sheets[sheetNameEvol]) {
        showLoading('Extrayendo las últimas 7 semanas de la hoja "EVOLUTIVO"...');
        const wsEvol = workbook.Sheets[sheetNameEvol];
        evolSecos = extractEvolutivoData(wsEvol, 5, 9);    // C5:BC9 para Secos (WHSE 655)
        evolFrescos = extractEvolutivoData(wsEvol, 14, 18); // C14:BC18 para Frescos (WHSE 676)
      }

      // 3. Procesar datos de Secos (655) y Frescos (676) desde la tabla tbBD
      showLoading('Calculando tablas dinámicas, zonas y Top SKUs para Secos y Frescos...');
      ACTIVE_DATABASE['655'] = compileWarehouseData(rawRows, '655', 'CD Secos 655', evolSecos);
      ACTIVE_DATABASE['676'] = compileWarehouseData(rawRows, '676', 'CD Frescos 676', evolFrescos);

      // 4. Actualizar estado visual
      document.getElementById('dropzoneBox').style.display = 'none';
      document.getElementById('fileStatusBar').style.display = 'flex';
      document.getElementById('loadedFileName').textContent = `Archivo: ${file.name}`;
      document.getElementById('loadedFileMeta').textContent = `${rawRows.length.toLocaleString('en-US')} filas procesadas de la hoja "${sheetNameBD}" | Hoja EVOLUTIVO conectada`;

      renderWarehouseData(currentWarehouse);
      hideLoading();
      showToast(`¡Archivo ${file.name} procesado con éxito!`, 'success', 4000);

    } catch (err) {
      console.error('Error al procesar Excel:', err);
      hideLoading();
      alert(`Error al leer el archivo Excel: ${err.message}\nVerifica que contenga las columnas especificadas.`);
    }
  };

  reader.onerror = function () {
    hideLoading();
    showToast('Error al leer el archivo del disco', 'danger');
  };

  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════════════════════════════════════════
// 3. EXTRACCIÓN DINÁMICA DE LA HOJA "EVOLUTIVO" (ÚLTIMAS 7 SEMANAS)
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Lee un rango horizontal de evolutivo (ej. C5:BC9) y extrae dinámicamente
 * las últimas 7 columnas que contengan datos.
 */
function extractEvolutivoData(ws, startRow1, endRow1) {
  // Convertir filas base 1 a base 0
  const rHeader = startRow1 - 1;
  const rDataStart = startRow1;
  const rDataEnd = endRow1 - 1;

  // Rango de columnas C a BC (C=2, BC=54)
  const colStart = 2; // 'C'
  const colEnd = 54;   // 'BC'

  // 1. Detectar columnas válidas que tengan datos en las filas de porcentajes
  const validCols = [];
  for (let c = colStart; c <= colEnd; c++) {
    const headerCell = ws[XLSX.utils.encode_cell({ r: rHeader, c: c })];
    const firstDataCell = ws[XLSX.utils.encode_cell({ r: rDataStart, c: c })];

    if (headerCell && headerCell.v !== undefined && headerCell.v !== '' &&
        firstDataCell && firstDataCell.v !== undefined && firstDataCell.v !== '') {
      validCols.push(c);
    }
  }

  if (validCols.length === 0) return null;

  // Tomar las últimas 7 columnas con datos
  const last7Cols = validCols.slice(-7);

  // Extraer nombres de semanas de los encabezados
  const weeks = last7Cols.map(c => {
    const cell = ws[XLSX.utils.encode_cell({ r: rHeader, c: c })];
    let txt = String(cell ? cell.v : '').trim();
    if (!txt.toUpperCase().startsWith('S-')) txt = 'S-' + txt;
    return txt;
  });

  // Etiquetas de los 4 rangos de antigüedad
  const rangeLabels = [
    { label: '0 a 10 Semanas', color: '#10b981' },
    { label: '10 a 25 Semanas', color: '#f59e0b' },
    { label: '25 a 52 Semanas', color: '#f97316' },
    { label: 'Mayor a 52 Semanas', color: '#ef4444' }
  ];

  // Extraer valores de porcentajes para cada fila
  const evolution = rangeLabels.map((rInfo, idx) => {
    const rowIdx = rDataStart + idx;
    const values = last7Cols.map(c => {
      const cell = ws[XLSX.utils.encode_cell({ r: rowIdx, c: c })];
      if (!cell || cell.v === undefined) return 0;
      let val = Number(cell.v);
      if (isNaN(val)) {
        val = parseFloat(String(cell.v).replace('%', '').replace(',', '.')) || 0;
      } else if (val <= 1.0) {
        val = val * 100; // Si viene en formato 0.8861 -> 88.61%
      }
      return parseFloat(val.toFixed(2));
    });
    return {
      label: rInfo.label,
      values: values,
      color: rInfo.color
    };
  });

  return { weeks, evolution };
}

// ══════════════════════════════════════════════════════════════════════════════
// 4. CÁLCULO Y COMPILACIÓN DE DATOS DESDE LA TABLA "tbBD"
// ══════════════════════════════════════════════════════════════════════════════

function normalizeRange(str) {
  if (!str) return '0 a 10 Semanas';
  const s = String(str).toLowerCase().trim();
  if (s.includes('0 a 10') || s.includes('0-10') || s.includes('< 10')) return '0 a 10 Semanas';
  if (s.includes('10 a 25') || s.includes('10-25')) return '10 a 25 Semanas';
  if (s.includes('25 a 52') || s.includes('25-52')) return '25 a 52 Semanas';
  if (s.includes('52') || s.includes('año') || s.includes('ano')) return 'mayor a 52 Semanas';
  return '0 a 10 Semanas';
}

function parseNum(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  const clean = String(v).replace(/,/g, '').trim();
  const n = parseFloat(clean);
  return isNaN(n) ? 0 : n;
}

function compileWarehouseData(rawRows, whseTarget, whseLabel, evolObj) {
  // Filtrar filas del almacén
  const rows = rawRows.filter(r => {
    const w = String(r['WHSE'] || r['whse'] || r['Whse'] || '').trim();
    return w === String(whseTarget);
  });

  if (rows.length === 0) {
    // Si no hay filas de ese WHSE en el archivo, devolver estructura base
    return createEmptyWarehouseData(whseTarget, whseLabel);
  }

  let totalCost = 0;
  let totalBultos = 0;
  const lpnsSet = new Set();

  // Rangos de semanas acumuladores
  const rangeAgg = {
    '0 a 10 Semanas': { cost: 0, bultos: 0, lpns: new Set(), color: '#10b981', status: 'Saludable' },
    '10 a 25 Semanas': { cost: 0, bultos: 0, lpns: new Set(), color: '#f59e0b', status: 'En Alerta' },
    '25 a 52 Semanas': { cost: 0, bultos: 0, lpns: new Set(), color: '#f97316', status: 'Riesgo Medio' },
    'mayor a 52 Semanas': { cost: 0, bultos: 0, lpns: new Set(), color: '#ef4444', status: 'Crítico >1 año' }
  };

  // Divisiones acumuladores
  const divisionAgg = {};

  // Zonas acumuladores (RCK, RHB, otros)
  const zoneAgg = {
    rck: { totalCost: 0, lpns: new Set(), bultos: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, skus52: {} },
    rhb: { totalCost: 0, lpns: new Set(), bultos: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, skus52: {} }
  };

  // SKUs >52 Semanas acumuladores
  const skusOver52 = {};

  // Ubicaciones físicas únicas acumuladores (Division x Rango)
  const locationsAgg = {};

  rows.forEach(r => {
    const cost = parseNum(r['COSTOS'] || r['Costos'] || r['costos'] || 0);
    const bultos = parseNum(r['BULTOS'] || r['Bultos'] || r['bultos'] || 0);
    const onHand = parseNum(r['ON_HAND'] || r['On_Hand'] || r['on_hand'] || 0);
    const dias = parseNum(r['DIAS'] || r['Dias'] || r['dias'] || 0);

    const lpn = String(r['LPN'] || r['lpn'] || Math.random()).trim();
    const sku = String(r['SKU'] || r['sku'] || '').trim();
    const desc = String(r['DESCRIPCION'] || r['Descripcion'] || r['descripcion'] || '').trim();
    const div = String(r['DIVISION'] || r['Division'] || r['division'] || 'OTROS').trim();
    const rng = normalizeRange(r['SEMANAS'] || r['Semanas'] || r['semanas']);
    const zona = String(r['ZONA'] || r['Zona'] || r['zona'] || '').toUpperCase().trim();
    const ubic = String(r['UBICACION'] || r['Ubicacion'] || r['ubicacion'] || '').trim();
    const rngFv = String(r['RNG-FV'] || r['rng-fv'] || r['FECHA_VENCIM'] || '-').trim();

    totalCost += cost;
    totalBultos += bultos;
    lpnsSet.add(lpn);

    // Acumular en Rangos
    if (rangeAgg[rng]) {
      rangeAgg[rng].cost += cost;
      rangeAgg[rng].bultos += bultos;
      rangeAgg[rng].lpns.add(lpn);
    }

    // Acumular en Divisiones
    if (!divisionAgg[div]) {
      divisionAgg[div] = { code: div, r010: 0, r1025: 0, r2552: 0, r52: 0, total: 0 };
    }
    divisionAgg[div].total += cost;
    if (rng === '0 a 10 Semanas') divisionAgg[div].r010 += cost;
    else if (rng === '10 a 25 Semanas') divisionAgg[div].r1025 += cost;
    else if (rng === '25 a 52 Semanas') divisionAgg[div].r2552 += cost;
    else if (rng === 'mayor a 52 Semanas') divisionAgg[div].r52 += cost;

    // Acumular en Zonas (RCK vs RHB)
    const isRck = zona.includes('RCK') || zona.includes('RACK');
    const isRhb = zona.includes('RHB') || zona.includes('HIGHBAY');
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

    // Top SKUs >52 Semanas global
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

    // Ubicaciones físicas únicas
    if (ubic) {
      if (!locationsAgg[div]) {
        locationsAgg[div] = {
          div,
          r010Ubic: new Set(),
          r1025Ubic: new Set(),
          r2552Ubic: new Set(),
          r52Ubic: new Set(),
          totalUbic: new Set()
        };
      }
      locationsAgg[div].totalUbic.add(ubic);
      if (rng === '0 a 10 Semanas') locationsAgg[div].r010Ubic.add(ubic);
      else if (rng === '10 a 25 Semanas') locationsAgg[div].r1025Ubic.add(ubic);
      else if (rng === '25 a 52 Semanas') locationsAgg[div].r2552Ubic.add(ubic);
      else if (rng === 'mayor a 52 Semanas') locationsAgg[div].r52Ubic.add(ubic);
    }
  });

  // Estructurar Rangos
  const ranges = Object.keys(rangeAgg).map(k => {
    const item = rangeAgg[k];
    return {
      label: k,
      cost: item.cost,
      costPct: totalCost > 0 ? (item.cost / totalCost) * 100 : 0,
      lpns: item.lpns.size,
      lpnsPct: lpnsSet.size > 0 ? (item.lpns.size / lpnsSet.size) * 100 : 0,
      bultos: item.bultos,
      bultosPct: totalBultos > 0 ? (item.bultos / totalBultos) * 100 : 0,
      color: item.color,
      status: item.status
    };
  });

  // Estructurar Divisiones ordenadas por costo total descendente
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

  // Estructurar Top 10 SKUs >52 Semanas
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

  // Estructurar Zonas
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

  // Estructurar Ubicaciones Físicas
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

  // Usar evolutivo extraído de la hoja "EVOLUTIVO" si existe, sino valores calculados
  const weeks = evolObj ? evolObj.weeks : ['S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36'];
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

function createEmptyWarehouseData(whseTarget, whseLabel) {
  return {
    name: whseLabel,
    whseCode: whseTarget,
    badgeText: whseTarget === '655' ? 'CD SECOS 655' : 'CD FRESCOS 676',
    totalCost: 0,
    totalLpns: 0,
    totalBultos: 0,
    ranges: [
      { label: '0 a 10 Semanas', cost: 0, costPct: 0, lpns: 0, bultos: 0, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 0, costPct: 0, lpns: 0, bultos: 0, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 0, costPct: 0, lpns: 0, bultos: 0, color: '#f97316', status: 'Riesgo Medio' },
      { label: 'mayor a 52 Semanas', cost: 0, costPct: 0, lpns: 0, bultos: 0, color: '#ef4444', status: 'Crítico >1 año' }
    ],
    divisions: [],
    top10Skus: [],
    zones: {
      rck: { totalCost: 0, lpns: 0, bultos: 0, pct: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, lpns52: 0, topSkus: [] },
      rhb: { totalCost: 0, lpns: 0, bultos: 0, pct: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, lpns52: 0, topSkus: [] }
    },
    weeks: ['S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36'],
    evolution: [],
    locationsTotal: 0,
    locations: []
  };
}

// ══════════════════════════════════════════════════════════════════════════════
// 5. DATOS DE EJEMPLO DE LA SEMANA 36 (OPCIÓN DE PREVISUALIZACIÓN)
// ══════════════════════════════════════════════════════════════════════════════

function loadSampleData() {
  showLoading('Cargando datos oficiales de ejemplo de la Semana 36...');

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
      { code: 'J10-BAZAR', r010: 2213932, r1025: 1180948, r2552: 486869, r52: 7664, total: 3889141, pct: 3.4 },
      { code: 'J06-PANADERIA Y PASTELERIA', r010: 1074602, r1025: 90837, r2552: 21380, r52: 1405, total: 1188224, pct: 1.1 },
      { code: 'J12-INSTITUCIONALES', r010: 887262, r1025: 0, r2552: 4370, r52: 872, total: 892505, pct: 0.8 },
      { code: 'J05-FLC', r010: 879846, r1025: 2994, r2552: 0, r52: 0, total: 882840, pct: 0.8 },
      { code: 'J07-PLATOS PREPARADOS', r010: 112807, r1025: 4247, r2552: 10184, r52: 22463, total: 149701, pct: 0.1 }
    ],
    top10Skus: [
      { sku: '43111173', desc: 'DURAZNO EN MITADES PRECIO UNO 415G', div: 'J01-PGC', rngFv: '19/07/2027 - 20/07/2027', lpns: 20, cost: 61831, bultos: 1110, onHand: 26640, days: 682, badge: 'CRÍTICO #1' },
      { sku: '43438965', desc: 'VINO TINTO ALBACORA X750ML', div: 'J01-PGC', rngFv: '12/12/2026 - 10/10/2028', lpns: 4, cost: 54671, bultos: 323, onHand: 1938, days: 550, badge: 'CRÍTICO #2' },
      { sku: '43491112', desc: 'COMBO MUG APILABLE VERANO PU', div: 'J09-HOGAR', rngFv: '-', lpns: 7, cost: 26720, bultos: 213, onHand: 213, days: 569, badge: 'TEMPORADA' },
      { sku: '42464523', desc: 'PULPA FINAL DE TOMATE TOTTUS X 400 G', div: 'J01-PGC', rngFv: '30/09/2027 - 30/09/2027', lpns: 6, cost: 26206, bultos: 785, onHand: 9420, days: 368, badge: 'MARCA PROP' },
      { sku: '43488563', desc: 'COMBO GUANTE CON SILIC Y TELA 2025', div: 'J09-HOGAR', rngFv: '-', lpns: 5, cost: 17027, bultos: 117, onHand: 117, days: 512, badge: 'TEMPORADA' },
      { sku: '41843146', desc: 'KETCHUP AMERICANO TOTTUS X 425GR', div: 'J01-PGC', rngFv: '03/03/2027 - 19/05/2027', lpns: 3, cost: 12653, bultos: 198, onHand: 3168, days: 473, badge: 'MARCA PROP' },
      { sku: '43439251', desc: 'LAMINA DE AJI S IMPRES 280MM PET PE 60U', div: 'J07-PLATOS', rngFv: '05/08/2026 - 05/08/2026', lpns: 1, cost: 11671, bultos: 55, onHand: 55, days: 395, badge: 'INSUMO' },
      { sku: '43439250', desc: 'LAMINA DE AJI S IMPRES 170MM PET PE 60U', div: 'J07-PLATOS', rngFv: '05/08/2026 - 05/08/2026', lpns: 1, cost: 10792, bultos: 73, onHand: 73, days: 395, badge: 'INSUMO' },
      { sku: '43491111', desc: 'COMBO MUG APILABLE VERANO CJ', div: 'J09-HOGAR', rngFv: '-', lpns: 3, cost: 10663, bultos: 85, onHand: 85, days: 569, badge: 'TEMPORADA' },
      { sku: '43314915', desc: 'SET X 2 ESPECIERO TAPA CORCHO 90ML', div: 'J09-HOGAR', rngFv: '-', lpns: 5, cost: 10227, bultos: 138, onHand: 3312, days: 492, badge: 'SALDOS' }
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
          { sku: '43111173', desc: 'DURAZNO EN MITADES PRECIO UNO', fv: '20/07/2027', lpns: 20, cost: 61831, bultos: 1110 },
          { sku: '43438965', desc: 'VINO TINTO ALBACORA X750ML', fv: '10/10/2028', lpns: 4, cost: 54671, bultos: 323 },
          { sku: '42464523', desc: 'PULPA FINAL DE TOMATE TOTTUS', fv: '30/09/2027', lpns: 6, cost: 26206, bultos: 785 },
          { sku: '43439251', desc: 'LAMINA AJI 280MM PET', fv: '05/08/2026', lpns: 1, cost: 11671, bultos: 55 },
          { sku: '43439250', desc: 'LAMINA AJI 170MM PET', fv: '05/08/2026', lpns: 1, cost: 10792, bultos: 73 },
          { sku: '42261627', desc: 'GANCHOS ROPA NEGRO X40', fv: '-', lpns: 1, cost: 5567, bultos: 81 },
          { sku: '43214510', desc: 'CUCHARA FIDEOS PLAST', fv: '-', lpns: 1, cost: 2704, bultos: 43 },
          { sku: '42261622', desc: 'LIGAS MIX TOTTUS X300', fv: '-', lpns: 1, cost: 2212, bultos: 16 },
          { sku: '41880323', desc: 'CAJA KEKE PREMIUM', fv: '-', lpns: 1, cost: 1405, bultos: 6 },
          { sku: '43121313', desc: 'JUEGO COMEDOR VIDRIO 4S', fv: '-', lpns: 1, cost: 1207, bultos: 4 }
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
          { sku: '43488563', desc: 'COMBO GUANTE SILIC Y TELA', fv: '-', lpns: 5, cost: 17027, bultos: 117 },
          { sku: '41843146', desc: 'KETCHUP AMERICANO TOTTUS', fv: '19/05/2027', lpns: 3, cost: 12653, bultos: 198 },
          { sku: '43491111', desc: 'COMBO MUG APILABLE CJ', fv: '-', lpns: 3, cost: 10663, bultos: 85 },
          { sku: '43314915', desc: 'SET X 2 ESPECIERO TAPA CORCHO', fv: '-', lpns: 5, cost: 10227, bultos: 138 },
          { sku: '43278081', desc: 'JUEGO SABANAS 1.5PLZ BLA', fv: '-', lpns: 3, cost: 7163, bultos: 67 },
          { sku: '42039096', desc: 'VINO MARQUES VITORIA BLANCO', fv: '-', lpns: 1, cost: 7016, bultos: 93 },
          { sku: '42039097', desc: 'VINO MARQUES VITORIA JOVEN', fv: '-', lpns: 1, cost: 4877, bultos: 53 },
          { sku: '43324607', desc: 'GARDEN PLAYHOUSE CON REJA', fv: '-', lpns: 3, cost: 3878, bultos: 18 },
          { sku: '43497327', desc: 'ALCANCIA CUPCAKE', fv: '-', lpns: 2, cost: 3020, bultos: 32 }
        ]
      }
    },
    weeks: ['S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36'],
    evolution: [
      { label: '0 a 10 Semanas', values: [88.61, 89.14, 87.07, 85.83, 86.07, 87.09, 87.96], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [9.47, 8.73, 10.68, 11.83, 11.71, 10.74, 10.18], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [1.60, 1.83, 1.95, 2.03, 1.94, 1.91, 1.60], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.32, 0.30, 0.31, 0.31, 0.28, 0.26, 0.27], color: '#ef4444' }
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

  ACTIVE_DATABASE['676'] = {
    name: 'CD Villa El Salvador - Frescos',
    whseCode: '676',
    badgeText: 'CD FRESCOS 676',
    totalCost: 38450210,
    totalLpns: 11240,
    totalBultos: 486200,
    ranges: [
      { label: '0 a 10 Semanas', cost: 35912496, costPct: 93.40, lpns: 10453, lpnsPct: 93.00, bultos: 457028, bultosPct: 94.00, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 2230112, costPct: 5.80, lpns: 685, lpnsPct: 6.09, bultos: 26255, bultosPct: 5.40, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 269151, costPct: 0.70, lpns: 86, lpnsPct: 0.77, bultos: 2431, bultosPct: 0.50, color: '#f97316', status: 'Riesgo Alto' },
      { label: 'mayor a 52 Semanas', cost: 38451, costPct: 0.10, lpns: 16, lpnsPct: 0.14, bultos: 486, bultosPct: 0.10, color: '#ef4444', status: 'Crítico Insumos' }
    ],
    divisions: [
      { code: 'J05-FLC (LÁCTEOS Y CONGELADOS)', r010: 16500000, r1025: 820000, r2552: 60000, r52: 8000, total: 17388000, pct: 45.2 },
      { code: 'J03-CARNES Y PESCADOS', r010: 9800000, r1025: 350000, r2552: 25000, r52: 0, total: 10175000, pct: 26.5 },
      { code: 'J04-FRUTAS Y VERDURAS', r010: 6100000, r1025: 120000, r2552: 0, r52: 0, total: 6220000, pct: 16.2 },
      { code: 'J06-PANADERIA Y PASTELERIA', r010: 2100000, r1025: 180000, r2552: 45000, r52: 2400, total: 2327400, pct: 6.1 },
      { code: 'J07-PLATOS PREPARADOS', r010: 950000, r1025: 65000, r2552: 12000, r52: 1800, total: 1028800, pct: 2.7 },
      { code: 'J01-PGC FRESCOS / INSUMOS', r010: 462496, r1025: 695112, r2552: 127151, r52: 26251, total: 1310910, pct: 3.4 }
    ],
    top10Skus: [
      { sku: '51200981', desc: 'EMPAQUE PET TERMOSELLABLE 500G', div: 'J07-PLATOS', rngFv: '01/01/2028', lpns: 4, cost: 14200, bultos: 180, onHand: 9000, days: 520, badge: 'INSUMO' },
      { sku: '50123984', desc: 'CAJA CONGELADOS POLIETILENO', div: 'J05-FLC', rngFv: '-', lpns: 3, cost: 9500, bultos: 120, onHand: 2400, days: 480, badge: 'MATERIAL' },
      { sku: '50983211', desc: 'PREMEZCLA PANETON INDUSTRIAL 25KG', div: 'J06-PANAD', rngFv: '15/10/2026', lpns: 2, cost: 4850, bultos: 50, onHand: 50, days: 420, badge: 'MAT. PRIMA' },
      { sku: '51456201', desc: 'ETIQUETA BALANZA TERMICA 58X43', div: 'J04-FRUT', rngFv: '-', lpns: 3, cost: 3800, bultos: 60, onHand: 600, days: 390, badge: 'SUMINISTRO' }
    ],
    zones: {
      rck: {
        totalCost: 15380084,
        lpns: 4496,
        bultos: 194480,
        pct: 40.0,
        r010: 14300000,
        r1025: 980000,
        r2552: 75000,
        r52: 25084,
        lpns52: 10,
        topSkus: [
          { sku: '51200981', desc: 'EMPAQUE PET TERMOSELLABLE 500G', fv: '01/01/2028', lpns: 4, cost: 14200, bultos: 180 },
          { sku: '50983211', desc: 'PREMEZCLA PANETON INDUSTRIAL', fv: '15/10/2026', lpns: 2, cost: 4850, bultos: 50 }
        ]
      },
      rhb: {
        totalCost: 23070126,
        lpns: 6744,
        bultos: 291720,
        pct: 60.0,
        r010: 21612496,
        r1025: 1250112,
        r2552: 194151,
        r52: 13367,
        lpns52: 6,
        topSkus: [
          { sku: '50123984', desc: 'CAJA CONGELADOS POLIETILENO', fv: '-', lpns: 3, cost: 9500, bultos: 120 }
        ]
      }
    },
    weeks: ['S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36'],
    evolution: [
      { label: '0 a 10 Semanas', values: [92.10, 92.80, 93.00, 91.50, 92.40, 93.10, 93.40], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [6.50, 5.90, 5.80, 7.10, 6.40, 5.90, 5.80], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [1.10, 1.00, 0.90, 1.20, 1.00, 0.80, 0.70], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.30, 0.30, 0.30, 0.20, 0.20, 0.20, 0.10], color: '#ef4444' }
    ],
    locationsTotal: 9850,
    locations: [
      { div: 'J05-FLC (CONGELADOS / LÁCTEOS)', r010: 4200, r1025: 280, r2552: 18, r52: 4, total: 4502, pct: 45.7 },
      { div: 'J03-CARNES Y PESCADOS', r010: 2400, r1025: 110, r2552: 12, r52: 0, total: 2522, pct: 25.6 }
    ]
  };

  document.getElementById('dropzoneBox').style.display = 'none';
  document.getElementById('fileStatusBar').style.display = 'flex';
  document.getElementById('loadedFileName').textContent = 'Datos de Ejemplo: Semana 36 (Oficiales)';
  document.getElementById('loadedFileMeta').textContent = 'Previsualización activa para CD Secos 655 y CD Frescos 676';

  renderWarehouseData(currentWarehouse);
  hideLoading();
  showToast('Datos oficiales de la Semana 36 cargados como ejemplo', 'info', 3000);
}

// ══════════════════════════════════════════════════════════════════════════════
// 6. RENDERIZADO VISUAL DEL DASHBOARD
// ══════════════════════════════════════════════════════════════════════════════

function switchWarehouse(whseCode) {
  currentWarehouse = whseCode;
  
  document.getElementById('btnWhseSecos').classList.toggle('active', whseCode === '655');
  document.getElementById('btnWhseFrescos').classList.toggle('active', whseCode === '676');
  
  renderWarehouseData(whseCode);
  showToast(`Mostrando datos de ${whseCode === '655' ? 'CD Secos (WHSE 655)' : 'CD Frescos (WHSE 676)'}`, 'info');
}

function switchSlide(slideNum) {
  currentSlide = slideNum;
  
  [1, 2, 3].forEach(n => {
    document.getElementById(`tabBtnSlide${n}`).classList.toggle('active', n === slideNum);
    const container = document.getElementById(`slide${n}-container`);
    if (container) {
      container.style.display = (n === slideNum) ? 'block' : 'none';
    }
  });

  setTimeout(() => {
    if (slideNum === 1 && chartInstanceS1) chartInstanceS1.resize();
    if (slideNum === 3 && chartInstanceS3) chartInstanceS3.resize();
  }, 60);
}

function renderWarehouseData(whseCode) {
  const data = ACTIVE_DATABASE[whseCode];
  if (!data) return;

  // Actualizar badges
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

  const r010 = data.ranges[0] || { costPct: 0, cost: 0, bultosPct: 0 };
  document.getElementById('s1-kpi-healthy-pct').textContent = `${r010.costPct.toFixed(1)}%`;
  document.getElementById('s1-kpi-healthy-cost').textContent = formatCurrency(r010.cost);
  document.getElementById('s1-kpi-healthy-bultos').textContent = `${r010.bultosPct.toFixed(1)}% Bultos`;

  const r1025 = data.ranges[1] || { costPct: 0, cost: 0, lpns: 0 };
  document.getElementById('s1-kpi-warn-pct').textContent = `${r1025.costPct.toFixed(1)}%`;
  document.getElementById('s1-kpi-warn-cost').textContent = formatCurrency(r1025.cost);
  document.getElementById('s1-kpi-warn-lpns').textContent = `${formatNumber(r1025.lpns)} LPNs`;

  const r2552 = data.ranges[2] || { cost: 0, costPct: 0 };
  const r52 = data.ranges[3] || { cost: 0, costPct: 0, lpns: 0 };
  const critCost = r2552.cost + r52.cost;
  const critPct = r2552.costPct + r52.costPct;
  document.getElementById('s1-kpi-crit-cost').textContent = formatCurrency(critCost);
  document.getElementById('s1-kpi-crit-pct').textContent = `${critPct.toFixed(1)}% del Capital`;
  document.getElementById('s1-kpi-over52').textContent = `>52s: ${formatCompact(r52.cost)} (${r52.lpns} LPNs)`;

  // Gráfico S1
  renderChartS1(data.ranges);

  // Tabla S1: Matriz Divisiones
  const tbodyDiv = document.getElementById('tbodyS1Divisions');
  tbodyDiv.innerHTML = '';
  data.divisions.forEach(d => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${d.code}</td>
      <td>${formatCompact(d.r010)}</td>
      <td>${formatCompact(d.r1025)}</td>
      <td style="${d.r2552 > 200000 ? 'background:#fffbeb; font-weight:700; color:#b45309;' : ''}">${formatCompact(d.r2552)}</td>
      <td style="${d.r52 > 10000 ? 'background:#fee2e2; font-weight:800; color:#dc2626;' : ''}">${d.r52 > 0 ? formatNumber(d.r52) : '-'}</td>
      <td style="font-weight:700;">${formatCompact(d.total)}</td>
      <td style="color:#64748b; font-weight:600;">${d.pct}%</td>
    `;
    tbodyDiv.appendChild(tr);
  });

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

  if (data.top10Skus.length === 0) {
    tbodyTop.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:18px; color:#64748b;">No se registraron SKUs con más de 52 semanas de antigüedad</td></tr>`;
  } else {
    data.top10Skus.forEach((sku, idx) => {
      sumTopCost += sku.cost;
      sumTopLpns += sku.lpns;
      sumTopBultos += sku.bultos;
      sumTopOnHand += sku.onHand;

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

  const labels = ranges.map(r => r.label);
  const dataCosts = ranges.map(r => (r.cost / 1000000).toFixed(2));
  const colors = ranges.map(r => r.color);

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
            label: (item) => ` S/ ${item.raw} Millones (${ranges[item.dataIndex].costPct.toFixed(1)}%)`
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
  const z = data.zones;

  // Zona RCK
  document.getElementById('s2-rck-cost').textContent = formatCurrency(z.rck.totalCost);
  document.getElementById('s2-rck-sub').textContent = `${z.rck.pct}% del CD | ${formatNumber(z.rck.lpns)} LPNs`;
  document.getElementById('s2-rck-010').textContent = formatCompact(z.rck.r010);
  document.getElementById('s2-rck-1025').textContent = formatCompact(z.rck.r1025);
  document.getElementById('s2-rck-2552').textContent = formatCompact(z.rck.r2552);
  document.getElementById('s2-rck-52').textContent = formatNumber(z.rck.r52);

  // Zona RHB
  document.getElementById('s2-rhb-cost').textContent = formatCurrency(z.rhb.totalCost);
  document.getElementById('s2-rhb-sub').textContent = `${z.rhb.pct}% del CD | ${formatNumber(z.rhb.lpns)} LPNs`;
  document.getElementById('s2-rhb-010').textContent = formatCompact(z.rhb.r010);
  document.getElementById('s2-rhb-1025').textContent = formatCompact(z.rhb.r1025);
  document.getElementById('s2-rhb-2552').textContent = formatCompact(z.rhb.r2552);
  document.getElementById('s2-rhb-52').textContent = formatNumber(z.rhb.r52);

  // Tablas RCK
  const tbodyRck = document.getElementById('tbodyS2Rck');
  tbodyRck.innerHTML = '';
  let sumRckCost = 0, sumRckLpns = 0, sumRckBultos = 0;

  if (z.rck.topSkus.length === 0) {
    tbodyRck.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:12px; color:#64748b;">No hay SKUs >52s en RCK</td></tr>`;
  } else {
    z.rck.topSkus.forEach(s => {
      sumRckCost += s.cost;
      sumRckLpns += s.lpns;
      sumRckBultos += s.bultos;
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

  // Tablas RHB
  const tbodyRhb = document.getElementById('tbodyS2Rhb');
  tbodyRhb.innerHTML = '';
  let sumRhbCost = 0, sumRhbLpns = 0, sumRhbBultos = 0;

  if (z.rhb.topSkus.length === 0) {
    tbodyRhb.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:12px; color:#64748b;">No hay SKUs >52s en RHB</td></tr>`;
  } else {
    z.rhb.topSkus.forEach(s => {
      sumRhbCost += s.cost;
      sumRhbLpns += s.lpns;
      sumRhbBultos += s.bultos;
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
  data.weeks.forEach((w, idx) => {
    const isCurrent = idx === data.weeks.length - 1;
    theadRow.innerHTML += `<th style="${isCurrent ? 'background:#eff6ff; color:#1d4ed8; font-weight:800;' : ''}">${w}</th>`;
  });
  theadRow.innerHTML += `<th>Var. WoW</th>`;

  const tbodyEvol = document.getElementById('tbodyS3Evolution');
  tbodyEvol.innerHTML = '';
  data.evolution.forEach(e => {
    const vals = e.values;
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
      trHtml += `<td style="${isCurrent ? 'background:#eff6ff; font-weight:800; color:' + e.color + ';' : ''}">${v.toFixed(2)}%</td>`;
    });

    trHtml += `
      <td style="font-weight:700; color:${isPositiveGood ? '#059669' : '#dc2626'};">
        ${diff >= 0 ? '▲ +' : '▼ '}${diff.toFixed(2)}%
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

  data.locations.forEach(loc => {
    sum010 += loc.r010;
    sum1025 += loc.r1025;
    sum2552 += loc.r2552;
    sum52 += loc.r52;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight:600;">${loc.div}</td>
      <td>${formatNumber(loc.r010)}</td>
      <td style="${loc.r1025 > 500 ? 'background:#fffbeb; font-weight:700; color:#b45309;' : ''}">${formatNumber(loc.r1025)}</td>
      <td style="${loc.r2552 > 100 ? 'background:#ffedd5; font-weight:700; color:#c2410c;' : ''}">${formatNumber(loc.r2552)}</td>
      <td style="${loc.r52 > 10 ? 'background:#fee2e2; font-weight:800; color:#dc2626;' : ''}">${loc.r52 > 0 ? loc.r52 : '-'}</td>
      <td style="font-weight:700;">${formatNumber(loc.total)}</td>
      <td style="color:#64748b; font-weight:600;">${loc.pct}%</td>
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

  const datasets = evolution.map(e => ({
    label: e.label,
    data: e.values,
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
      labels: weeks,
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
            label: (item) => ` ${item.dataset.label}: ${item.raw.toFixed(2)}%`
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
// 7. MOTOR DE CAPTURA 16:9 PARA POWERPOINT
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
      if (!blob) throw new Error('Error al generar blob');

      if (navigator.clipboard && window.ClipboardItem) {
        try {
          const item = new ClipboardItem({ 'image/png': blob });
          await navigator.clipboard.write([item]);
          showToast(`✅ ¡Lámina ${currentSlide} copiada en 16:9! Lista para pegar en PowerPoint (Ctrl + V)`, 'success', 5000);
        } catch (clipErr) {
          console.warn('Error al copiar al portapapeles:', clipErr);
          downloadCanvasAsPng(offscreen, `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`);
          showToast(`Descargada como PNG 16:9 (Permiso de portapapeles restringido)`, 'info', 4500);
        }
      } else {
        downloadCanvasAsPng(offscreen, `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`);
        showToast(`Descargada como PNG 16:9`, 'info', 4500);
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

// ══════════════════════════════════════════════════════════════════════════════
// 8. FORMATO Y UTILIDADES
// ══════════════════════════════════════════════════════════════════════════════

function formatCurrency(val) {
  if (val === null || val === undefined) return 'S/ 0';
  return 'S/ ' + Math.round(val).toLocaleString('en-US');
}

function formatNumber(val) {
  if (val === null || val === undefined) return '0';
  return Math.round(val).toLocaleString('en-US');
}

function formatCompact(val) {
  if (!val) return '0';
  if (val >= 1000000) {
    return (val / 1000000).toFixed(2) + 'M';
  }
  if (val >= 1000) {
    return (val / 1000).toFixed(1) + 'K';
  }
  return val.toString();
}

function showToast(message, type = 'info', duration = 3500) {
  const toast = document.getElementById('toastMessage');
  const toastText = document.getElementById('toastText');
  if (!toast || !toastText) return;

  toastText.textContent = message;
  
  if (type === 'success') {
    toast.style.background = '#065f46';
  } else if (type === 'danger') {
    toast.style.background = '#991b1b';
  } else {
    toast.style.background = '#0f172a';
  }

  toast.classList.add('show');
  setTimeout(() => {
    toast.classList.remove('show');
  }, duration);
}
