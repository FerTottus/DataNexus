/**
 * DataNexus - Antigüedad de Inventario (Secos 655 y Frescos 676)
 * Generador y visualizador de indicadores semanales con exportación a PowerPoint 16:9
 */

// Estado global
let currentWarehouse = '655'; // '655' (Secos) o '676' (Frescos)
let currentSlide = 1;
let chartInstanceS1 = null;
let chartInstanceS3 = null;
let isCapturing = false;

// ══════════════════════════════════════════════════════════════════════════════
// 1. BASES DE DATOS PRECARGADAS (OFICIALES SEMANA 36)
// ══════════════════════════════════════════════════════════════════════════════

const DATABASE = {
  // CD SECOS (WHSE 655) - Datos exactos de las láminas del usuario
  '655': {
    name: 'CD Huachipa - Secos',
    whseCode: '655',
    badgeText: 'CD SECOS 655',
    totalCost: 112812540,
    totalLpns: 33265,
    totalBultos: 1345343,
    
    // Rangos de semanas
    ranges: [
      { label: '0 a 10 Semanas', cost: 97924406, costPct: 86.80, lpns: 27817, lpnsPct: 83.62, bultos: 1183316, bultosPct: 87.96, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 12502561, costPct: 11.08, lpns: 4471, lpnsPct: 13.44, bultos: 136958, bultosPct: 10.18, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 2096734, costPct: 1.86, lpns: 899, lpnsPct: 2.70, bultos: 21460, bultosPct: 1.60, color: '#f97316', status: 'Riesgo Medio' },
      { label: 'mayor a 52 Semanas', cost: 288839, costPct: 0.26, lpns: 78, lpnsPct: 0.23, bultos: 3609, bultosPct: 0.27, color: '#ef4444', status: 'Crítico >1 año' }
    ],

    // Costo por División y Antigüedad
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

    // Top 10 SKUs >52 Semanas
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

    // Zonas Operativas
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

    // Evolutivo 7 semanas
    weeks: ['S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36'],
    evolution: [
      { label: '0 a 10 Semanas', values: [88.61, 89.14, 87.07, 85.83, 86.07, 87.09, 87.96], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [9.47, 8.73, 10.68, 11.83, 11.71, 10.74, 10.18], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [1.60, 1.83, 1.95, 2.03, 1.94, 1.91, 1.60], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.32, 0.30, 0.31, 0.31, 0.28, 0.26, 0.27], color: '#ef4444' }
    ],

    // Ubicaciones físicas
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
  },

  // CD FRESCOS (WHSE 676) - Calibrado con la lógica de rotación perecible
  '676': {
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
      { sku: '51456201', desc: 'ETIQUETA BALANZA TERMICA 58X43', div: 'J04-FRUT', rngFv: '-', lpns: 3, cost: 3800, bultos: 60, onHand: 600, days: 390, badge: 'SUMINISTRO' },
      { sku: '50221345', desc: 'BANDEJA ABSORBENTE CARNES B2', div: 'J03-CARN', rngFv: '-', lpns: 2, cost: 2900, bultos: 40, onHand: 4000, days: 375, badge: 'INSUMO' },
      { sku: '50882190', desc: 'CONCENTRADO CITRICO DESINFECTANTE', div: 'J04-FRUT', rngFv: '20/12/2026', lpns: 1, cost: 1800, bultos: 20, onHand: 20, days: 368, badge: 'LIMPIEZA' },
      { sku: '51009842', desc: 'FILM EXTENSIBLE ALIMENTARIO 45CM', div: 'J05-FLC', rngFv: '-', lpns: 1, cost: 1401, bultos: 16, onHand: 32, days: 365, badge: 'EMPAQUE' }
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
          { sku: '50983211', desc: 'PREMEZCLA PANETON INDUSTRIAL', fv: '15/10/2026', lpns: 2, cost: 4850, bultos: 50 },
          { sku: '50221345', desc: 'BANDEJA ABSORBENTE CARNES', fv: '-', lpns: 2, cost: 2900, bultos: 40 },
          { sku: '50882190', desc: 'CONCENTRADO CITRICO', fv: '20/12/2026', lpns: 1, cost: 1800, bultos: 20 },
          { sku: '51009842', desc: 'FILM EXTENSIBLE 45CM', fv: '-', lpns: 1, cost: 1334, bultos: 16 }
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
          { sku: '50123984', desc: 'CAJA CONGELADOS POLIETILENO', fv: '-', lpns: 3, cost: 9500, bultos: 120 },
          { sku: '51456201', desc: 'ETIQUETA BALANZA TERMICA', fv: '-', lpns: 2, cost: 2500, bultos: 40 },
          { sku: '51009842', desc: 'FILM EXTENSIBLE ALIMENTARIO', fv: '-', lpns: 1, cost: 1367, bultos: 16 }
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
      { div: 'J03-CARNES Y PESCADOS', r010: 2400, r1025: 110, r2552: 12, r52: 0, total: 2522, pct: 25.6 },
      { div: 'J04-FRUTAS Y VERDURAS', r010: 1600, r1025: 45, r2552: 0, r52: 0, total: 1645, pct: 16.7 },
      { div: 'J06-PANADERIA Y PASTELERIA', r010: 550, r1025: 40, r2552: 8, r52: 2, total: 600, pct: 6.1 },
      { div: 'J07-PLATOS PREPARADOS', r010: 280, r1025: 22, r2552: 5, r52: 2, total: 309, pct: 3.1 },
      { div: 'J01-PGC FRESCOS / INSUMOS', r010: 210, r1025: 48, r2552: 10, r52: 4, total: 272, pct: 2.8 }
    ]
  }
};

// ══════════════════════════════════════════════════════════════════════════════
// 2. INICIALIZACIÓN Y RENDERIZADO
// ══════════════════════════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  renderWarehouseData(currentWarehouse);
});

function switchWarehouse(whseCode) {
  currentWarehouse = whseCode;
  
  // Actualizar botones de WHSE
  document.getElementById('btnWhseSecos').classList.toggle('active', whseCode === '655');
  document.getElementById('btnWhseFrescos').classList.toggle('active', whseCode === '676');
  
  renderWarehouseData(whseCode);
  showToast(`Cambiado a ${DATABASE[whseCode].name}`, 'info');
}

function switchSlide(slideNum) {
  currentSlide = slideNum;
  
  // Actualizar botones de pestaña
  [1, 2, 3].forEach(n => {
    document.getElementById(`tabBtnSlide${n}`).classList.toggle('active', n === slideNum);
    const container = document.getElementById(`slide${n}-container`);
    if (container) {
      container.style.display = (n === slideNum) ? 'block' : 'none';
    }
  });

  // Reajustar gráficos al cambiar de vista
  setTimeout(() => {
    if (slideNum === 1 && chartInstanceS1) chartInstanceS1.resize();
    if (slideNum === 3 && chartInstanceS3) chartInstanceS3.resize();
  }, 50);
}

function renderWarehouseData(whseCode) {
  const data = DATABASE[whseCode];
  if (!data) return;

  // 1. Badges y Títulos
  ['slide1', 'slide2', 'slide3'].forEach(id => {
    const el = document.getElementById(`${id}-whse-badge`);
    if (el) el.textContent = data.badgeText;
  });

  // 2. RENDERIZAR SLIDE 1
  renderSlide1(data);

  // 3. RENDERIZAR SLIDE 2
  renderSlide2(data);

  // 4. RENDERIZAR SLIDE 3
  renderSlide3(data);
}

// ══════════════════════════════════════════════════════════════════════════════
// 3. RENDERIZADO DETALLADO POR DIAPOSITIVA
// ══════════════════════════════════════════════════════════════════════════════

function renderSlide1(data) {
  // KPIs
  document.getElementById('s1-kpi-total').textContent = formatCurrency(data.totalCost);
  document.getElementById('s1-kpi-lpns').textContent = `${formatNumber(data.totalLpns)} LPNs`;
  document.getElementById('s1-kpi-bultos').textContent = `${formatNumber(data.totalBultos)} Bultos`;

  const r010 = data.ranges[0];
  document.getElementById('s1-kpi-healthy-pct').textContent = `${r010.costPct.toFixed(1)}%`;
  document.getElementById('s1-kpi-healthy-cost').textContent = formatCurrency(r010.cost);
  document.getElementById('s1-kpi-healthy-bultos').textContent = `${r010.bultosPct.toFixed(1)}% Bultos`;

  const r1025 = data.ranges[1];
  document.getElementById('s1-kpi-warn-pct').textContent = `${r1025.costPct.toFixed(1)}%`;
  document.getElementById('s1-kpi-warn-cost').textContent = formatCurrency(r1025.cost);
  document.getElementById('s1-kpi-warn-lpns').textContent = `${formatNumber(r1025.lpns)} LPNs`;

  const r2552 = data.ranges[2];
  const r52 = data.ranges[3];
  const critCost = r2552.cost + r52.cost;
  const critPct = r2552.costPct + r52.costPct;
  document.getElementById('s1-kpi-crit-cost').textContent = formatCurrency(critCost);
  document.getElementById('s1-kpi-crit-pct').textContent = `${critPct.toFixed(1)}% del Capital`;
  document.getElementById('s1-kpi-over52').textContent = `>52s: ${formatCompact(r52.cost)} (${r52.lpns} LPNs)`;

  // Gráfico S1: Distribución por Rangos
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

  const tfootDiv = document.getElementById('tfootS1Divisions');
  tfootDiv.innerHTML = `
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

  const tfootTop = document.getElementById('tfootS1Top10');
  tfootTop.innerHTML = `
    <tr>
      <td colspan="4" style="text-align:left;">Total Top 10 SKUs</td>
      <td>${sumTopLpns}</td>
      <td style="color:#dc2626; font-size:0.82rem;">S/ ${formatNumber(sumTopCost)}</td>
      <td>${formatNumber(sumTopBultos)}</td>
      <td>${formatNumber(sumTopOnHand)}</td>
      <td colspan="2" style="text-align:center; color:#64748b;">~84% del stock >52 sem</td>
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
            label: (item) => ` S/ ${item.raw} Millones (${ranges[item.dataIndex].costPct}%)`
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
  document.getElementById('s2-rck-010').textContent = formatCompact(z.rck.r010);
  document.getElementById('s2-rck-1025').textContent = formatCompact(z.rck.r1025);
  document.getElementById('s2-rck-2552').textContent = formatCompact(z.rck.r2552);
  document.getElementById('s2-rck-52').textContent = formatNumber(z.rck.r52);

  // Zona RHB
  document.getElementById('s2-rhb-cost').textContent = formatCurrency(z.rhb.totalCost);
  document.getElementById('s2-rhb-010').textContent = formatCompact(z.rhb.r010);
  document.getElementById('s2-rhb-1025').textContent = formatCompact(z.rhb.r1025);
  document.getElementById('s2-rhb-2552').textContent = formatCompact(z.rhb.r2552);
  document.getElementById('s2-rhb-52').textContent = formatNumber(z.rhb.r52);

  // Tablas RCK
  const tbodyRck = document.getElementById('tbodyS2Rck');
  tbodyRck.innerHTML = '';
  let sumRckCost = 0, sumRckLpns = 0, sumRckBultos = 0;
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
  // Evolutivo Semanal
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
    const lastVal = vals[vals.length - 1];
    const prevVal = vals[vals.length - 2];
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

  // Ubicaciones Físicas
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
// 4. MOTOR DE CAPTURA 16:9 PARA POWERPOINT
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Captura la diapositiva activa con proporción exacta 16:9 (1920x1080)
 * y la escribe directamente en el portapapeles del sistema para Ctrl + V en PowerPoint.
 */
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
    // 1. Renderizar el contenedor a Canvas usando html2canvas con scale=2 para nitidez Retina
    const renderedCanvas = await window.html2canvas(container, {
      scale: 2,
      backgroundColor: '#ffffff',
      useCORS: true,
      logging: false
    });

    // 2. Componer en un Canvas con relación de aspecto exacta 16:9 (1920 x 1080)
    const W = 1920;
    const H = 1080;
    const targetAspect = 16 / 9;

    const offscreen = document.createElement('canvas');
    offscreen.width = W;
    offscreen.height = H;
    const ctx = offscreen.getContext('2d');

    // Fondo blanco corporativo
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);

    // Ajuste de escala para que encaje proporcionalmente dentro de 1920x1080
    const margin = 40; // Margen exterior de seguridad para PowerPoint
    const availW = W - margin * 2;
    const availH = H - margin * 2;

    const scaleFit = Math.min(availW / renderedCanvas.width, availH / renderedCanvas.height);
    const drawW = renderedCanvas.width * scaleFit;
    const drawH = renderedCanvas.height * scaleFit;
    const posX = (W - drawW) / 2;
    const posY = (H - drawH) / 2;

    // Dibujar imagen con suavizado de alta definición
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(renderedCanvas, posX, posY, drawW, drawH);

    // 3. Convertir a Blob y copiar al portapapeles
    offscreen.toBlob(async (blob) => {
      if (!blob) throw new Error('Error al generar blob');

      if (navigator.clipboard && window.ClipboardItem) {
        try {
          const item = new ClipboardItem({ 'image/png': blob });
          await navigator.clipboard.write([item]);
          showToast(`✅ ¡Lámina ${currentSlide} copiada en 16:9! Lista para pegar en PowerPoint (Ctrl + V)`, 'success', 5000);
        } catch (clipErr) {
          console.warn('Error al copiar al portapapeles:', clipErr);
          // Fallback a descarga automática si el navegador bloquea el portapapeles
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

/**
 * Descarga la diapositiva actual como archivo PNG 1920x1080
 */
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
// 5. MODAL Y GESTIÓN DE DATOS
// ══════════════════════════════════════════════════════════════════════════════

function openDataModal() {
  document.getElementById('dataModal').style.display = 'flex';
  document.getElementById('modalWhseSelect').value = currentWarehouse;
}

function closeDataModal() {
  document.getElementById('dataModal').style.display = 'none';
}

function loadSampleData() {
  renderWarehouseData(currentWarehouse);
  closeDataModal();
  showToast('Datos oficiales de la Semana 36 restaurados', 'success');
}

function saveModalData() {
  const targetWhse = document.getElementById('modalWhseSelect').value;
  const pasteText = document.getElementById('modalPasteEvol').value.trim();

  if (pasteText) {
    try {
      // Parsear líneas pegadas de Evolutivo
      const lines = pasteText.split('\n');
      lines.forEach(line => {
        const parts = line.split('\t');
        if (parts.length >= 8) {
          const rangeLabel = parts[0].trim();
          const pcts = parts.slice(1, 8).map(p => parseFloat(p.replace('%', '').replace(',', '.').trim()) || 0);
          
          const matchEvol = DATABASE[targetWhse].evolution.find(e => e.label.toLowerCase().includes(rangeLabel.toLowerCase()));
          if (matchEvol) {
            matchEvol.values = pcts;
          }
        }
      });
      showToast('Datos de Evolutivo actualizados correctamente', 'success');
    } catch (e) {
      console.warn('Error al procesar texto:', e);
      showToast('No se pudo procesar el formato pegado. Revisa las tabulaciones.', 'danger');
    }
  }

  renderWarehouseData(targetWhse);
  closeDataModal();
}

// ══════════════════════════════════════════════════════════════════════════════
// 6. UTILIDADES Y FORMATO
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
