/**
 * DataNexus - Días a Tienda (DTI) Dashboard Logic
 * Procesamiento de archivos de trabajo Excel para gráficos evolutivos
 */

// Registrar plugin de etiquetas para Chart.js
Chart.register(ChartDataLabels);

let rawWorkbook = null;
let currentDataset = [];
let chartInstances = {};

// Normalizador de texto para mapeo inteligente de encabezados
function normalizeHeader(str) {
  if (!str) return '';
  return str.toString()
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // sin acentos
    .replace(/[._\-\s]+/g, ' ')
    .trim();
}

// Inicialización
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  setupDragAndDrop();
});

function setupEventListeners() {
  const fileInput = document.getElementById('fileInput');
  fileInput.addEventListener('change', handleFileSelect);

  document.getElementById('sheetSelect').addEventListener('change', handleSheetChange);
  document.getElementById('filterCD').addEventListener('change', applyFiltersAndRender);
  document.getElementById('filterZona').addEventListener('change', (e) => {
    // Ajustar meta sugerida automáticamente
    const metaInput = document.getElementById('inputMeta');
    if (e.target.value === 'Local') metaInput.value = '1.5';
    else if (e.target.value === 'Provincia') metaInput.value = '2.5';
    applyFiltersAndRender();
  });
  document.getElementById('filterAnio').addEventListener('change', applyFiltersAndRender);
  document.getElementById('filterMetrica').addEventListener('change', applyFiltersAndRender);
  document.getElementById('inputMeta').addEventListener('input', applyFiltersAndRender);
  document.getElementById('btnResetFilters').addEventListener('click', resetFilters);
  document.getElementById('btnLoadDemo').addEventListener('click', loadDemoData);
  document.getElementById('btnExportCsv').addEventListener('click', exportSummaryToCsv);
}

function setupDragAndDrop() {
  const dropZone = document.getElementById('dropZone');

  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
    });
  });

  dropZone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      processFile(files[0]);
    }
  });
}

function handleFileSelect(e) {
  const file = e.target.files[0];
  if (file) {
    processFile(file);
  }
}

// Procesar el archivo con SheetJS
function processFile(file) {
  const reader = new FileReader();

  document.getElementById('fileNameDisplay').textContent = "Procesando " + file.name + "...";
  document.getElementById('fileBanner').classList.remove('hidden');

  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      rawWorkbook = XLSX.read(data, { type: 'array', cellDates: true });

      const sheetSelect = document.getElementById('sheetSelect');
      sheetSelect.innerHTML = '';

      let defaultSheet = rawWorkbook.SheetNames[0];

      // Priorizar hojas clave del tutorial: 'Data general', 'DATA', 'PICKING FOLIOS', 'Consolidado'
      const keySheets = ['data general', 'data', 'consolidado', 'picking folios', 'hoja1'];
      for (const name of rawWorkbook.SheetNames) {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        sheetSelect.appendChild(opt);

        const normName = normalizeHeader(name);
        for (const ks of keySheets) {
          if (normName.includes(ks)) {
            defaultSheet = name;
            break;
          }
        }
      }

      sheetSelect.value = defaultSheet;
      document.getElementById('fileNameDisplay').textContent = file.name;

      parseSheetData(defaultSheet);

    } catch (err) {
      console.error(err);
      alert("Error al leer el archivo Excel: " + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
}

function handleSheetChange(e) {
  if (!rawWorkbook) return;
  parseSheetData(e.target.value);
}

// Extraer y normalizar registros de la hoja seleccionada
function parseSheetData(sheetName) {
  const worksheet = rawWorkbook.Sheets[sheetName];
  if (!worksheet) return;

  const rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: null, raw: false });
  if (!rawJson || rawJson.length === 0) {
    alert("La pestaña seleccionada está vacía.");
    return;
  }

  // Detectar encabezados dinámicamente
  const sample = rawJson[0];
  const headerMap = mapHeaders(sample);

  currentDataset = [];

  rawJson.forEach(row => {
    // CD (568/655 -> 655 Secos, 569/676 -> 676 Frescos)
    let rawCd = headerMap.cd ? row[headerMap.cd] : '';
    let cdClean = '655';
    if (rawCd) {
      const strCd = rawCd.toString().trim();
      if (strCd.includes('569') || strCd.includes('676') || strCd.toLowerCase().includes('fresco')) {
        cdClean = '676';
      } else {
        cdClean = '655';
      }
    }

    // Zona (Local vs Provincia)
    let rawZona = headerMap.zona ? row[headerMap.zona] : '';
    let zonaClean = 'Local';
    if (rawZona) {
      const strZ = rawZona.toString().toLowerCase();
      if (strZ.includes('prov') || strZ === 'p') zonaClean = 'Provincia';
      else zonaClean = 'Local';
    }

    // Semana
    let rawSemana = headerMap.semana ? row[headerMap.semana] : null;
    let semNum = null;
    if (rawSemana !== null && rawSemana !== undefined) {
      const match = rawSemana.toString().match(/\d+/);
      if (match) semNum = parseInt(match[0], 10);
    }

    // Año
    let rawAnio = headerMap.anio ? row[headerMap.anio] : null;
    let anioNum = 2026;
    if (rawAnio) {
      const match = rawAnio.toString().match(/202[0-9]/);
      if (match) anioNum = parseInt(match[0], 10);
    }

    // Si viene fecha y no semana/año, calcularlos
    if (headerMap.fecha && row[headerMap.fecha] && (!semNum || !anioNum)) {
      const d = new Date(row[headerMap.fecha]);
      if (!isNaN(d.getTime())) {
        if (!anioNum) anioNum = d.getFullYear();
        if (!semNum) semNum = getWeekNumber(d);
      }
    }

    // Si aún no hay semana, asignar por defecto
    if (!semNum) semNum = 1;

    // Métricas numéricas
    const dtiVal = parseNumeric(headerMap.dti ? row[headerMap.dti] : null);
    const opeInternaVal = parseNumeric(headerMap.opeInterna ? row[headerMap.opeInterna] : null);
    const opeSalidaVal = parseNumeric(headerMap.opeSalida ? row[headerMap.opeSalida] : null);
    const cajasVal = parseNumeric(headerMap.cajas ? row[headerMap.cajas] : 1);

    if (dtiVal !== null && dtiVal >= 0) {
      currentDataset.push({
        cd: cdClean,
        zona: zonaClean,
        semana: semNum,
        anio: anioNum,
        dti: dtiVal,
        opeInterna: opeInternaVal !== null ? opeInternaVal : dtiVal * 0.45,
        opeSalida: opeSalidaVal !== null ? opeSalidaVal : dtiVal * 0.55,
        cajas: cajasVal && cajasVal > 0 ? cajasVal : 1
      });
    }
  });

  document.getElementById('fileRowsDisplay').textContent = `(${currentDataset.length.toLocaleString()} registros válidos)`;
  document.getElementById('emptyState').classList.add('hidden');
  document.getElementById('filtersSection').classList.remove('hidden');
  document.getElementById('dashboardContent').classList.remove('hidden');

  applyFiltersAndRender();
}

// Mapeo inteligente de encabezados
function mapHeaders(sampleRow) {
  const map = {
    cd: null,
    zona: null,
    semana: null,
    anio: null,
    fecha: null,
    dti: null,
    opeInterna: null,
    opeSalida: null,
    cajas: null
  };

  const keys = Object.keys(sampleRow);

  keys.forEach(k => {
    const n = normalizeHeader(k);

    // CD
    if (!map.cd && (n === 'cd' || n.includes('centro') || n.includes('cod cd') || n.includes('codigo cd'))) {
      map.cd = k;
    }
    // Zona
    if (!map.zona && (n.includes('tipo tienda') || n === 'tipo' || n === 'zona' || n.includes('destino') || n.includes('local provincia'))) {
      map.zona = k;
    }
    // Semana
    if (!map.semana && (n.includes('semana') || n === 'sem' || n === 'wk')) {
      map.semana = k;
    }
    // Año
    if (!map.anio && (n === 'ano' || n === 'anio' || n === 'year' || n === 'a o')) {
      map.anio = k;
    }
    // Fecha
    if (!map.fecha && (n.includes('fecha') || n.includes('date'))) {
      map.fecha = k;
    }
    // DTI
    if (!map.dti && (n.includes('dti final') || n === 'dti' || n === 'dit d' || n.includes('dias a tienda') || n.includes('dias tienda') || n === 'dit')) {
      map.dti = k;
    }
    // Operación Interna
    if (!map.opeInterna && (n.includes('ope interna') || n.includes('opeinterna') || n.includes('interna') || n.includes('picking a psl'))) {
      map.opeInterna = k;
    }
    // Operación Salida
    if (!map.opeSalida && (n.includes('ope salida') || n.includes('opesalida') || n.includes('salida') || n.includes('psl a cargas'))) {
      map.opeSalida = k;
    }
    // Cajas
    if (!map.cajas && (n.includes('caja') || n.includes('cantidad') || n.includes('bulto') || n.includes('unidades'))) {
      map.cajas = k;
    }
  });

  return map;
}

function parseNumeric(val) {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return val;
  const cleaned = val.toString().replace(/,/g, '.').trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}

function getWeekNumber(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}

// Aplicar Filtros y Calcular
function applyFiltersAndRender() {
  if (!currentDataset || currentDataset.length === 0) return;

  const cdFilter = document.getElementById('filterCD').value;
  const zonaFilter = document.getElementById('filterZona').value;
  const anioFilter = document.getElementById('filterAnio').value;
  const metricaFilter = document.getElementById('filterMetrica').value;
  const metaVal = parseFloat(document.getElementById('inputMeta').value) || 1.5;

  // Filtrar dataset
  let filtered = currentDataset.filter(item => {
    if (cdFilter !== 'ALL' && item.cd !== cdFilter) return false;
    if (zonaFilter !== 'ALL' && item.zona !== zonaFilter) return false;
    if (anioFilter !== 'ALL' && item.anio.toString() !== anioFilter) return false;
    return true;
  });

  if (filtered.length === 0) {
    alert("No se encontraron registros con los filtros seleccionados.");
    return;
  }

  // Agrupar por semana y año
  const weeklyAgg = {};

  filtered.forEach(row => {
    const key = `${row.anio}_S${row.semana.toString().padStart(2, '0')}`;
    if (!weeklyAgg[key]) {
      weeklyAgg[key] = {
        key: key,
        anio: row.anio,
        semana: row.semana,
        cd: row.cd,
        zona: row.zona,
        totalCajas: 0,
        sumDtiPonderado: 0,
        sumDtiSimple: 0,
        sumOpeInternaPonderada: 0,
        sumOpeSalidaPonderada: 0,
        count: 0
      };
    }

    weeklyAgg[key].totalCajas += row.cajas;
    weeklyAgg[key].sumDtiPonderado += (row.dti * row.cajas);
    weeklyAgg[key].sumDtiSimple += row.dti;
    weeklyAgg[key].sumOpeInternaPonderada += (row.opeInterna * row.cajas);
    weeklyAgg[key].sumOpeSalidaPonderada += (row.opeSalida * row.cajas);
    weeklyAgg[key].count += 1;
  });

  const weeklyList = Object.values(weeklyAgg).map(w => {
    const dtiFinal = metricaFilter === 'ponderado' 
      ? (w.totalCajas > 0 ? w.sumDtiPonderado / w.totalCajas : 0)
      : (w.count > 0 ? w.sumDtiSimple / w.count : 0);

    const opeIntFinal = w.totalCajas > 0 ? w.sumOpeInternaPonderada / w.totalCajas : 0;
    const opeSalFinal = w.totalCajas > 0 ? w.sumOpeSalidaPonderada / w.totalCajas : 0;

    return {
      ...w,
      dti: parseFloat(dtiFinal.toFixed(2)),
      opeInterna: parseFloat(opeIntFinal.toFixed(2)),
      opeSalida: parseFloat(opeSalFinal.toFixed(2))
    };
  }).sort((a, b) => {
    if (a.anio !== b.anio) return a.anio - b.anio;
    return a.semana - b.semana;
  });

  updateKpis(weeklyList, filtered, metaVal);
  renderCharts(weeklyList, metaVal, cdFilter, zonaFilter);
  renderSummaryTable(weeklyList, metaVal);
}

// Actualizar KPIs
function updateKpis(weeklyList, rawFiltered, metaVal) {
  if (weeklyList.length === 0) return;

  const lastWeek = weeklyList[weeklyList.length - 1];
  const dtiActual = lastWeek.dti;

  document.getElementById('kpiDtiActual').textContent = dtiActual.toFixed(2);
  document.getElementById('kpiTitle1').textContent = `DTI Sem ${lastWeek.semana} (${lastWeek.anio})`;

  const badgeSla = document.getElementById('kpiBadgeSla');
  const textSla = document.getElementById('kpiSlaText');

  if (dtiActual <= metaVal) {
    badgeSla.className = 'kpi-badge success';
    textSla.textContent = `Cumple Meta (≤ ${metaVal}d)`;
  } else {
    badgeSla.className = 'kpi-badge danger';
    textSla.textContent = `Excede Meta (${(dtiActual - metaVal).toFixed(2)}d)`;
  }

  // Promedio de todo el periodo
  const totalDti = weeklyList.reduce((acc, w) => acc + w.dti, 0);
  const avgDti = totalDti / weeklyList.length;
  document.getElementById('kpiDtiPromedio').textContent = avgDti.toFixed(2);
  document.getElementById('kpiSemanasCount').textContent = `${weeklyList.length} semanas analizadas`;

  // Desglose
  document.getElementById('kpiDesgloseTiempos').textContent = `Int: ${lastWeek.opeInterna}d | Sal: ${lastWeek.opeSalida}d`;

  // Volumen
  const totalCajas = rawFiltered.reduce((acc, r) => acc + r.cajas, 0);
  document.getElementById('kpiTotalCajas').textContent = Math.round(totalCajas).toLocaleString();
  document.getElementById('kpiTotalRegistros').textContent = `${rawFiltered.length.toLocaleString()} despachos`;
}

// Renderizar Gráficos con Chart.js
function renderCharts(weeklyList, metaVal, cdFilter, zonaFilter) {
  renderEvolutivoChart(weeklyList, metaVal, cdFilter, zonaFilter);
  renderComparativoChart(weeklyList);
  renderLeadTimeChart(weeklyList);
  renderLocalProvinciaChart();
}

function renderEvolutivoChart(weeklyList, metaVal, cdFilter, zonaFilter) {
  const ctx = document.getElementById('chartDtiEvolutivo').getContext('2d');
  if (chartInstances.evolutivo) chartInstances.evolutivo.destroy();

  const labels = weeklyList.map(w => `Sem ${w.semana}`);
  const dataDti = weeklyList.map(w => w.dti);
  const dataMeta = weeklyList.map(() => metaVal);

  const pointColors = dataDti.map(val => val <= metaVal ? '#10b981' : '#ef4444');

  let titleCD = cdFilter === '655' ? 'CD Secos (655)' : (cdFilter === '676' ? 'CD Frescos (676)' : 'Todos los CDs');
  document.getElementById('chart1Title').textContent = `Evolutivo Semanal DTI - ${titleCD} (${zonaFilter})`;

  chartInstances.evolutivo = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'DTI Real (Días)',
          data: dataDti,
          borderColor: '#2563eb',
          backgroundColor: '#3b82f6',
          pointBackgroundColor: pointColors,
          pointBorderColor: '#ffffff',
          pointBorderWidth: 2,
          pointRadius: 6,
          pointHoverRadius: 8,
          borderWidth: 2.5,
          tension: 0.25,
          datalabels: {
            align: 'top',
            anchor: 'end',
            color: '#0f172a',
            font: { weight: 'bold', size: 10 },
            formatter: (v) => v.toFixed(2)
          }
        },
        {
          label: `Meta SLA (${metaVal} días)`,
          data: dataMeta,
          borderColor: '#ef4444',
          borderDash: [5, 5],
          borderWidth: 2,
          pointRadius: 0,
          fill: false,
          datalabels: { display: false }
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.raw} días`
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          title: { display: true, text: 'Días a Tienda (DTI)' }
        },
        x: {
          grid: { display: false }
        }
      }
    }
  });
}

function renderComparativoChart(weeklyList) {
  const ctx = document.getElementById('chartDtiComparativo').getContext('2d');
  if (chartInstances.comparativo) chartInstances.comparativo.destroy();

  // Filtrar 2025 y 2026
  const data2025 = {};
  const data2026 = {};

  currentDataset.forEach(row => {
    if (row.anio === 2025) {
      if (!data2025[row.semana]) data2025[row.semana] = { sum: 0, c: 0 };
      data2025[row.semana].sum += row.dti;
      data2025[row.semana].c += 1;
    } else if (row.anio === 2026) {
      if (!data2026[row.semana]) data2026[row.semana] = { sum: 0, c: 0 };
      data2026[row.semana].sum += row.dti;
      data2026[row.semana].c += 1;
    }
  });

  const allWeeks = Array.from(new Set([...Object.keys(data2025), ...Object.keys(data2026)]))
    .map(Number).sort((a, b) => a - b);

  const labels = allWeeks.map(w => `Sem ${w}`);
  const vals2025 = allWeeks.map(w => data2025[w] ? parseFloat((data2025[w].sum / data2025[w].c).toFixed(2)) : null);
  const vals2026 = allWeeks.map(w => data2026[w] ? parseFloat((data2026[w].sum / data2026[w].c).toFixed(2)) : null);

  chartInstances.comparativo = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Año 2025',
          data: vals2025,
          backgroundColor: '#94a3b8',
          borderRadius: 4,
          datalabels: { display: false }
        },
        {
          label: 'Año 2026',
          data: vals2026,
          backgroundColor: '#2563eb',
          borderRadius: 4,
          datalabels: {
            align: 'top',
            anchor: 'end',
            font: { size: 9, weight: 'bold' },
            formatter: (v) => v ? v.toFixed(2) : ''
          }
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' }
      },
      scales: {
        y: { beginAtZero: true, title: { display: true, text: 'DTI (Días)' } },
        x: { grid: { display: false } }
      }
    }
  });
}

function renderLeadTimeChart(weeklyList) {
  const ctx = document.getElementById('chartLeadTime').getContext('2d');
  if (chartInstances.leadTime) chartInstances.leadTime.destroy();

  const labels = weeklyList.map(w => `Sem ${w.semana}`);
  const dataInterna = weeklyList.map(w => w.opeInterna);
  const dataSalida = weeklyList.map(w => w.opeSalida);

  chartInstances.leadTime = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Operación Interna (Picking -> PSL)',
          data: dataInterna,
          backgroundColor: '#3b82f6',
          stack: 'leadtime',
          borderRadius: 2,
          datalabels: { display: false }
        },
        {
          label: 'Operación Salida (PSL -> Cargas)',
          data: dataSalida,
          backgroundColor: '#f97316',
          stack: 'leadtime',
          borderRadius: 4,
          datalabels: {
            align: 'top',
            anchor: 'end',
            font: { size: 9, weight: 'bold' },
            formatter: (v, ctx) => {
              const total = dataInterna[ctx.dataIndex] + v;
              return total.toFixed(2);
            }
          }
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' }
      },
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, beginAtZero: true, title: { display: true, text: 'Días Operativos' } }
      }
    }
  });
}

function renderLocalProvinciaChart() {
  const ctx = document.getElementById('chartLocalProvincia').getContext('2d');
  if (chartInstances.localProvincia) chartInstances.localProvincia.destroy();

  const localWeeks = {};
  const provWeeks = {};

  currentDataset.forEach(r => {
    if (r.zona === 'Local') {
      if (!localWeeks[r.semana]) localWeeks[r.semana] = { sum: 0, c: 0 };
      localWeeks[r.semana].sum += r.dti;
      localWeeks[r.semana].c += 1;
    } else {
      if (!provWeeks[r.semana]) provWeeks[r.semana] = { sum: 0, c: 0 };
      provWeeks[r.semana].sum += r.dti;
      provWeeks[r.semana].c += 1;
    }
  });

  const allWeeks = Array.from(new Set([...Object.keys(localWeeks), ...Object.keys(provWeeks)]))
    .map(Number).sort((a, b) => a - b);

  const labels = allWeeks.map(w => `Sem ${w}`);
  const dataLocal = allWeeks.map(w => localWeeks[w] ? parseFloat((localWeeks[w].sum / localWeeks[w].c).toFixed(2)) : null);
  const dataProv = allWeeks.map(w => provWeeks[w] ? parseFloat((provWeeks[w].sum / provWeeks[w].c).toFixed(2)) : null);

  chartInstances.localProvincia = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Local (Meta 1.5d)',
          data: dataLocal,
          borderColor: '#0ea5e9',
          backgroundColor: '#0ea5e9',
          borderWidth: 2,
          pointRadius: 4,
          tension: 0.25,
          datalabels: { display: false }
        },
        {
          label: 'Provincia (Meta 2.5d)',
          data: dataProv,
          borderColor: '#8b5cf6',
          backgroundColor: '#8b5cf6',
          borderWidth: 2,
          pointRadius: 4,
          tension: 0.25,
          datalabels: { display: false }
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' }
      },
      scales: {
        y: { beginAtZero: true, title: { display: true, text: 'DTI (Días)' } },
        x: { grid: { display: false } }
      }
    }
  });
}

// Renderizar Tabla Resumen
function renderSummaryTable(weeklyList, metaVal) {
  const tbody = document.getElementById('summaryTableBody');
  tbody.innerHTML = '';

  weeklyList.forEach(w => {
    const tr = document.createElement('tr');
    const cumple = w.dti <= metaVal;

    tr.innerHTML = `
      <td><strong>Semana ${w.semana}</strong></td>
      <td>${w.anio}</td>
      <td>${w.cd === '655' ? 'CD 655 (Secos)' : 'CD 676 (Frescos)'}</td>
      <td>${w.zona}</td>
      <td class="text-right">${Math.round(w.totalCajas).toLocaleString()}</td>
      <td class="text-right">${w.opeInterna.toFixed(2)} d</td>
      <td class="text-right">${w.opeSalida.toFixed(2)} d</td>
      <td class="text-right"><strong>${w.dti.toFixed(2)} d</strong></td>
      <td class="text-right">${metaVal.toFixed(1)} d</td>
      <td class="text-center">
        <span class="cell-pill ${cumple ? 'success' : 'danger'}">
          ${cumple ? 'CUMPLE' : 'DESVÍO'}
        </span>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// Copiar imagen del gráfico para pegar en PowerPoint
function copyChartImage(canvasId) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  canvas.toBlob(blob => {
    const item = new ClipboardItem({ 'image/png': blob });
    navigator.clipboard.write([item]).then(() => {
      alert("¡Gráfico copiado al portapapeles! Ya puedes presionar Ctrl + V en PowerPoint.");
    }).catch(err => {
      console.error(err);
      downloadChartImage(canvasId, 'Grafico_DTI');
    });
  });
}

// Descargar imagen PNG
function downloadChartImage(canvasId, fileName) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const link = document.createElement('a');
  link.download = `${fileName}.png`;
  link.href = canvas.toDataURL('image/png');
  link.click();
}

// Exportar resumen a CSV
function exportSummaryToCsv() {
  const table = document.getElementById('summaryTable');
  let csv = [];
  for (let row of table.rows) {
    let cols = [];
    for (let cell of row.cells) {
      cols.push('"' + cell.innerText.replace(/"/g, '""').trim() + '"');
    }
    csv.push(cols.join(';'));
  }

  const blob = new Blob(["\uFEFF" + csv.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Resumen_Semanas_DTI.csv';
  a.click();
}

function resetFilters() {
  document.getElementById('filterCD').value = '655';
  document.getElementById('filterZona').value = 'Local';
  document.getElementById('filterAnio').value = '2026';
  document.getElementById('filterMetrica').value = 'ponderado';
  document.getElementById('inputMeta').value = '1.5';
  applyFiltersAndRender();
}

// Cargar Datos Demo representativos del tutorial de Falabella/Tottus
function loadDemoData() {
  const demoWeeks = [];
  const years = [2025, 2026];

  years.forEach(yr => {
    const maxSem = yr === 2025 ? 52 : 37;
    for (let s = 1; s <= maxSem; s++) {
      // Secos Local (Meta 1.5)
      demoWeeks.push({
        cd: '655', zona: 'Local', semana: s, anio: yr,
        dti: parseFloat((1.20 + Math.sin(s * 0.4) * 0.35 + (yr === 2025 ? 0.25 : 0)).toFixed(2)),
        opeInterna: parseFloat((0.65 + Math.sin(s * 0.3) * 0.15).toFixed(2)),
        opeSalida: parseFloat((0.65 + Math.cos(s * 0.3) * 0.15).toFixed(2)),
        cajas: Math.floor(45000 + Math.random() * 15000)
      });
      // Secos Provincia (Meta 2.5)
      demoWeeks.push({
        cd: '655', zona: 'Provincia', semana: s, anio: yr,
        dti: parseFloat((2.10 + Math.cos(s * 0.4) * 0.45 + (yr === 2025 ? 0.30 : 0)).toFixed(2)),
        opeInterna: parseFloat((0.95 + Math.sin(s * 0.3) * 0.20).toFixed(2)),
        opeSalida: parseFloat((1.30 + Math.cos(s * 0.3) * 0.20).toFixed(2)),
        cajas: Math.floor(25000 + Math.random() * 8000)
      });
      // Frescos Local (Meta 1.5)
      demoWeeks.push({
        cd: '676', zona: 'Local', semana: s, anio: yr,
        dti: parseFloat((1.10 + Math.sin(s * 0.5) * 0.25).toFixed(2)),
        opeInterna: parseFloat((0.55 + Math.sin(s * 0.2) * 0.10).toFixed(2)),
        opeSalida: parseFloat((0.55 + Math.cos(s * 0.2) * 0.10).toFixed(2)),
        cajas: Math.floor(30000 + Math.random() * 10000)
      });
      // Frescos Provincia (Meta 2.5)
      demoWeeks.push({
        cd: '676', zona: 'Provincia', semana: s, anio: yr,
        dti: parseFloat((2.05 + Math.cos(s * 0.5) * 0.35).toFixed(2)),
        opeInterna: parseFloat((0.85 + Math.sin(s * 0.2) * 0.15).toFixed(2)),
        opeSalida: parseFloat((1.20 + Math.cos(s * 0.2) * 0.15).toFixed(2)),
        cajas: Math.floor(18000 + Math.random() * 6000)
      });
    }
  });

  currentDataset = demoWeeks;

  document.getElementById('fileNameDisplay').textContent = "Datos Demo Representativos (Falabella/Tottus Sem 1-37)";
  document.getElementById('fileRowsDisplay').textContent = `(${currentDataset.length} registros cargados)`;
  document.getElementById('fileBanner').classList.remove('hidden');
  document.getElementById('emptyState').classList.add('hidden');
  document.getElementById('filtersSection').classList.remove('hidden');
  document.getElementById('dashboardContent').classList.remove('hidden');

  applyFiltersAndRender();
}
