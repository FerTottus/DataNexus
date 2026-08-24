/**
 * Pivot Engine & Dynamic Table Builder for SheetPivot
 * Gestiona el cálculo de tablas dinámicas (agrupaciones, sumas, promedios, etc.),
 * tablas planas con filtros avanzados, y el renderizado HTML con capacidades de copiado.
 */

const PivotEngine = {
  /**
   * Aplica filtros a un conjunto de filas
   * @param {Array<Object>} rows - Filas de datos
   * @param {Array<{field: string, operator: string, value: string}>} filters - Filtros activos
   * @param {string} globalSearch - Término de búsqueda general
   * @returns {Array<Object>} Filas filtradas
   */
  applyFilters(rows, filters = [], globalSearch = '') {
    if (!rows || rows.length === 0) return [];

    let filtered = rows;

    // 1. Filtro global de búsqueda
    if (globalSearch && globalSearch.trim()) {
      const term = globalSearch.trim().toLowerCase();
      filtered = filtered.filter(row => {
        return Object.values(row).some(v => {
          if (v === null || v === undefined) return false;
          return String(v).toLowerCase().includes(term);
        });
      });
    }

    // 2. Filtros específicos por columna
    if (filters && filters.length > 0) {
      filtered = filtered.filter(row => {
        return filters.every(f => {
          if (!f.field || f.value === undefined || f.value === '') return true;
          const cellVal = row[f.field];
          const filterVal = String(f.value).toLowerCase();
          const targetStr = cellVal !== null && cellVal !== undefined ? String(cellVal).toLowerCase() : '';
          const numCell = Number(cellVal);
          const numFilter = Number(f.value);

          switch (f.operator) {
            case 'equals':
              return targetStr === filterVal;
            case 'contains':
              return targetStr.includes(filterVal);
            case 'starts_with':
              return targetStr.startsWith(filterVal);
            case 'not_equals':
              return targetStr !== filterVal;
            case 'gt': // Mayor que
              return !isNaN(numCell) && !isNaN(numFilter) ? numCell > numFilter : false;
            case 'gte': // Mayor o igual
              return !isNaN(numCell) && !isNaN(numFilter) ? numCell >= numFilter : false;
            case 'lt': // Menor que
              return !isNaN(numCell) && !isNaN(numFilter) ? numCell < numFilter : false;
            case 'lte': // Menor o igual
              return !isNaN(numCell) && !isNaN(numFilter) ? numCell <= numFilter : false;
            default:
              return targetStr.includes(filterVal);
          }
        });
      });
    }

    return filtered;
  },

  /**
   * Realiza la agregación de una lista de valores según la función seleccionada
   * @param {Array<any>} values 
   * @param {'SUM'|'COUNT'|'AVG'|'MIN'|'MAX'} aggFunc 
   * @returns {number}
   */
  calculateAggregation(values, aggFunc = 'SUM') {
    if (!values || values.length === 0) return 0;

    if (aggFunc === 'COUNT') {
      return values.filter(v => v !== null && v !== undefined && v !== '').length;
    }

    // Convertir a numéricos válidos
    const numbers = values
      .map(v => {
        if (typeof v === 'number') return v;
        if (typeof v === 'string') {
          const cleaned = v.replace(/[^0-9.-]+/g, '');
          const n = parseFloat(cleaned);
          return isNaN(n) ? null : n;
        }
        return null;
      })
      .filter(v => v !== null && !isNaN(v));

    if (numbers.length === 0) return 0;

    switch (aggFunc) {
      case 'SUM':
        return numbers.reduce((acc, cur) => acc + cur, 0);
      case 'AVG':
        return numbers.reduce((acc, cur) => acc + cur, 0) / numbers.length;
      case 'MIN':
        return Math.min(...numbers);
      case 'MAX':
        return Math.max(...numbers);
      default:
        return numbers.reduce((acc, cur) => acc + cur, 0);
    }
  },

  buildPivotData(rows, config) {
    const {
      rowFields = [],
      colFields = [],
      valField = '',
      aggFunc = 'SUM',
      filters = [],
      globalSearch = ''
    } = config;

    // 1. Aplicar filtros específicos a los datos BASE primero
    const filteredRows = this.applyFilters(rows, filters, '');

    if (rowFields.length === 0 && colFields.length === 0) {
      const allValues = filteredRows.map(r => r[valField]);
      const total = this.calculateAggregation(allValues, aggFunc);
      return {
        headers: ['Total General'],
        matrixRows: [[this.formatNumber(total)]],
        rowCount: filteredRows.length
      };
    }

    // Estructuras de datos para O(N)
    const pivotMatrix = new Map(); // rowKey -> Map(colKey -> Array<values>)
    const colSet = new Set();
    const rowTotalsMap = new Map(); // rowKey -> Array<values>
    const colTotalsMap = new Map(); // colKey -> Array<values>
    const allValues = [];

    // 2. ÚNICA PASADA: Agrupar datos en O(N)
    for (let i = 0; i < filteredRows.length; i++) {
      const r = filteredRows[i];
      
      const rowKey = rowFields.length > 0 
        ? rowFields.map(rf => (r[rf] !== null && r[rf] !== undefined ? String(r[rf]) : '(Vacío)')).join(' | ')
        : 'Total';
        
      const colKey = colFields.length > 0
        ? colFields.map(cf => (r[cf] !== null && r[cf] !== undefined ? String(r[cf]) : '(Vacío)')).join(' - ')
        : 'Total';

      const val = r[valField];

      if (colFields.length > 0) colSet.add(colKey);

      // Guardar en matriz principal
      if (!pivotMatrix.has(rowKey)) {
        pivotMatrix.set(rowKey, new Map());
        rowTotalsMap.set(rowKey, []);
      }
      const rowMap = pivotMatrix.get(rowKey);
      if (!rowMap.has(colKey)) {
        rowMap.set(colKey, []);
      }
      rowMap.get(colKey).push(val);

      // Acumuladores de filas y columnas
      rowTotalsMap.get(rowKey).push(val);
      
      if (colFields.length > 0) {
        if (!colTotalsMap.has(colKey)) colTotalsMap.set(colKey, []);
        colTotalsMap.get(colKey).push(val);
      }
      
      allValues.push(val);
    }

    // 3. Preparar encabezados
    const pivotColValues = Array.from(colSet).sort();
    const headers = [...rowFields];
    if (pivotColValues.length > 0) {
      headers.push(...pivotColValues);
    }
    headers.push(`Total (${aggFunc})`);

    // 4. Construir matriz de salida
    const matrixRows = [];
    const sortedRowKeys = Array.from(pivotMatrix.keys()).sort();

    for (let i = 0; i < sortedRowKeys.length; i++) {
      const rowKey = sortedRowKeys[i];
      const rowMap = pivotMatrix.get(rowKey);
      const rowParts = rowFields.length > 0 ? rowKey.split(' | ') : [];
      const matrixRow = [...rowParts];

      if (pivotColValues.length > 0) {
        for (let j = 0; j < pivotColValues.length; j++) {
          const colKey = pivotColValues[j];
          const cellValues = rowMap.get(colKey) || [];
          const calculated = this.calculateAggregation(cellValues, aggFunc);
          matrixRow.push(this.formatNumber(calculated));
        }
      } else {
        // Sin columnas, la celda es el total de la fila
        const cellValues = rowMap.get('Total') || [];
        const calculated = this.calculateAggregation(cellValues, aggFunc);
        matrixRow.push(this.formatNumber(calculated));
      }

      // Añadir el total de la fila al final
      const rowTotalVals = rowTotalsMap.get(rowKey) || [];
      const rowTotal = this.calculateAggregation(rowTotalVals, aggFunc);
      matrixRow.push(this.formatNumber(rowTotal));

      matrixRows.push(matrixRow);
    }

    // 5. Fila de Totales Generales
    const grandTotalRow = new Array(rowFields.length).fill('');
    if (rowFields.length > 0) grandTotalRow[0] = 'Total General';

    if (pivotColValues.length > 0) {
      for (let j = 0; j < pivotColValues.length; j++) {
        const colKey = pivotColValues[j];
        const colVals = colTotalsMap.get(colKey) || [];
        const total = this.calculateAggregation(colVals, aggFunc);
        grandTotalRow.push(this.formatNumber(total));
      }
    }

    const grandTotal = this.calculateAggregation(allValues, aggFunc);
    grandTotalRow.push(this.formatNumber(grandTotal));

    // 6. Filtro Búsqueda Global (Aplicado a la matriz VISIBLE resultante)
    let finalMatrixRows = matrixRows;
    if (globalSearch && globalSearch.trim()) {
      const term = globalSearch.trim().toLowerCase();
      finalMatrixRows = finalMatrixRows.filter(rowCells => {
        return rowCells.some(cellVal => {
          if (cellVal === null || cellVal === undefined) return false;
          const cleanStr = String(cellVal).toLowerCase().replace(/,/g, '');
          const normalStr = String(cellVal).toLowerCase();
          return normalStr.includes(term) || cleanStr.includes(term);
        });
      });
    }

    return {
      headers,
      matrixRows: finalMatrixRows,
      grandTotalRow,
      rowCount: filteredRows.length
    };
  },

  /**
   * Construye los datos de una tabla plana / filtrada
   * @param {Array<Object>} rows 
   * @param {Object} config 
   * @returns {Object}
   */
  buildFlatData(rows, config) {
    const {
      selectedColumns = [],
      filters = [],
      globalSearch = '',
      sortBy = '',
      sortOrder = 'asc',
      limit = 200
    } = config;

    let filtered = this.applyFilters(rows, filters, globalSearch);

    // Ordenamiento
    if (sortBy) {
      filtered.sort((a, b) => {
        const valA = a[sortBy];
        const valB = b[sortBy];
        if (valA === valB) return 0;
        if (valA === null || valA === undefined) return 1;
        if (valB === null || valB === undefined) return -1;

        const numA = Number(valA);
        const numB = Number(valB);
        if (!isNaN(numA) && !isNaN(numB)) {
          return sortOrder === 'asc' ? numA - numB : numB - numA;
        }

        const comp = String(valA).localeCompare(String(valB), undefined, { numeric: true });
        return sortOrder === 'asc' ? comp : -comp;
      });
    }

    const totalCount = filtered.length;
    const paginated = limit > 0 ? filtered.slice(0, limit) : filtered;

    return {
      headers: selectedColumns,
      rows: paginated,
      totalCount
    };
  },

  /**
   * Formatea un número para visualización limpia
   * @param {number} num 
   * @returns {string}
   */
  formatNumber(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    // Si es entero, sin decimales
    if (Number.isInteger(num)) {
      return num.toLocaleString('en-US');
    }
    // Si tiene decimales, redondear a 2 dígitos
    return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  },

  /**
   * Genera el HTML de la tabla interactiva y configura los eventos de copiado
   * @param {Object} tableModel - { id, type, headers, rows/matrixRows, grandTotalRow }
   * @returns {string} HTML string
   */
  renderTableHtml(tableModel) {
    const { id, type, headers, matrixRows, rows, grandTotalRow, totalCount } = tableModel;

    let html = `<div class="table-responsive-wrapper" id="wrapper_${id}">`;
    html += `<table class="data-table" id="table_${id}">`;

    // Encabezados
    html += `<thead><tr>`;
    headers.forEach((h, colIdx) => {
      html += `
        <th data-col-index="${colIdx}" data-col-name="${h}">
          <div class="th-content">
            <span>${h}</span>
            <button class="btn-col-copy" title="Copiar columna ${h}" onclick="PivotEngine.handleCopyColumnClick('${id}', ${colIdx}, '${h}')">
              <i class="fa-solid fa-copy"></i>
            </button>
          </div>
        </th>
      `;
    });
    html += `</tr></thead>`;

    // Cuerpo
    html += `<tbody>`;

    if (type === 'pivot') {
      if (!matrixRows || matrixRows.length === 0) {
        html += `<tr><td colspan="${headers.length}" style="text-align: center; color: var(--text-muted); padding: 2rem;">No hay registros que coincidan con los filtros</td></tr>`;
      } else {
        matrixRows.forEach(rowCells => {
          html += `<tr>`;
          rowCells.forEach((cellVal, cIdx) => {
            const isNumeric = typeof cellVal === 'number' || (typeof cellVal === 'string' && /^-?[0-9,.]+$/.test(cellVal.trim()));
            const cssClass = isNumeric ? 'numeric' : '';
            html += `<td class="${cssClass}">${cellVal !== null && cellVal !== undefined ? cellVal : ''}</td>`;
          });
          html += `</tr>`;
        });

        // Total general
        if (grandTotalRow && grandTotalRow.length > 0) {
          html += `<tr class="total-row">`;
          grandTotalRow.forEach(val => {
            const isNumeric = typeof val === 'number' || (typeof val === 'string' && /^-?[0-9,.]+$/.test(val.trim()));
            const cssClass = isNumeric ? 'numeric' : '';
            html += `<td class="${cssClass}">${val}</td>`;
          });
          html += `</tr>`;
        }
      }
    } else {
      // Tipo 'flat'
      if (!rows || rows.length === 0) {
        html += `<tr><td colspan="${headers.length}" style="text-align: center; color: var(--text-muted); padding: 2rem;">No hay registros que coincidan con los filtros</td></tr>`;
      } else {
        rows.forEach(rowObj => {
          html += `<tr>`;
          headers.forEach(h => {
            const cellVal = rowObj[h];
            const isNumeric = typeof cellVal === 'number';
            const cssClass = isNumeric ? 'numeric' : '';
            // No formateamos números en tabla plana, los dejamos crudos como en la base
            const displayVal = (cellVal !== null && cellVal !== undefined) ? cellVal : '';
            html += `<td class="${cssClass}">${displayVal}</td>`;
          });
          html += `</tr>`;
        });
      }
    }

    html += `</tbody></table></div>`;
    return html;
  },

  /**
   * Callback para copiar una columna completa de una tabla específica
   */
  handleCopyColumnClick(tableId, colIndex, colName) {
    const tableEl = document.getElementById(`table_${tableId}`);
    if (!tableEl) return;

    const cells = Array.from(tableEl.querySelectorAll(`tbody tr:not(.total-row) td:nth-child(${colIndex + 1})`));
    const values = cells.map(td => td.innerText.trim());

    if (window.ClipboardUtil) {
      window.ClipboardUtil.copyColumn(colName, values);
    }
  }
};

window.PivotEngine = PivotEngine;
