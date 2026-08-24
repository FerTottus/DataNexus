/**
 * Cell Selection Script for SheetPivot
 * Enables Excel-like drag-to-select and copy for table cells.
 */

(function() {
  let isSelecting = false;
  let startCell = null;
  let currentEndCell = null;
  let selectionBounds = null;

  // Add styles dynamically
  const style = document.createElement('style');
  style.innerHTML = `
    .data-table td.cell-selected {
      background-color: rgba(37, 99, 235, 0.2) !important;
      outline: 1px solid rgba(37, 99, 235, 0.8);
      outline-offset: -1px;
    }
    .data-table {
      user-select: none; /* Previene la selección de texto nativa mientras arrastras */
    }
  `;
  document.head.appendChild(style);

  function getCellCoords(cell) {
    const tr = cell.parentElement;
    const tbody = tr.parentElement;
    const rowIndex = Array.prototype.indexOf.call(tbody.children, tr);
    const colIndex = Array.prototype.indexOf.call(tr.children, cell);
    return { row: rowIndex, col: colIndex, table: tr.closest('table') };
  }

  function clearSelection(table) {
    if (!table) return;
    const selected = table.querySelectorAll('.cell-selected');
    selected.forEach(c => c.classList.remove('cell-selected'));
    selectionBounds = null;
  }

  function renderSelection() {
    if (!startCell || !currentEndCell || startCell.table !== currentEndCell.table) return;

    clearSelection(startCell.table);

    const minRow = Math.min(startCell.row, currentEndCell.row);
    const maxRow = Math.max(startCell.row, currentEndCell.row);
    const minCol = Math.min(startCell.col, currentEndCell.col);
    const maxCol = Math.max(startCell.col, currentEndCell.col);

    selectionBounds = { minRow, maxRow, minCol, maxCol, table: startCell.table };

    const rows = startCell.table.querySelectorAll('tbody tr');
    for (let r = minRow; r <= maxRow; r++) {
      const tr = rows[r];
      if (!tr) continue;
      for (let c = minCol; c <= maxCol; c++) {
        const td = tr.children[c];
        if (td) td.classList.add('cell-selected');
      }
    }
  }

  document.addEventListener('mousedown', (e) => {
    const td = e.target.closest('td');
    if (td && td.closest('.data-table')) {
      if (e.button !== 0) return; // Only left click
      
      // Limpiamos selecciones previas en TODAS las tablas
      document.querySelectorAll('.data-table').forEach(t => clearSelection(t));

      isSelecting = true;
      startCell = getCellCoords(td);
      currentEndCell = startCell;
      renderSelection();
    } else if (!e.target.closest('.data-table')) {
      // Click fuera de las tablas, limpiamos
      document.querySelectorAll('.data-table').forEach(t => clearSelection(t));
      startCell = null;
      selectionBounds = null;
    }
  });

  document.addEventListener('mouseover', (e) => {
    if (!isSelecting) return;
    const td = e.target.closest('td');
    if (td && td.closest('.data-table')) {
      const coords = getCellCoords(td);
      if (coords.table === startCell.table) {
        currentEndCell = coords;
        renderSelection();
      }
    }
  });

  document.addEventListener('mouseup', () => {
    isSelecting = false;
  });

  // Interceptar Ctrl+C para copiar celdas seleccionadas
  document.addEventListener('copy', (e) => {
    if (selectionBounds && startCell) {
      // Tenemos celdas seleccionadas personalizadas, procedemos
      e.preventDefault();
      
      const { minRow, maxRow, minCol, maxCol, table } = selectionBounds;
      const rows = table.querySelectorAll('tbody tr');
      
      let tsvText = "";
      let htmlText = "<table><tbody>";

      for (let r = minRow; r <= maxRow; r++) {
        const tr = rows[r];
        if (!tr) continue;
        
        let rowText = [];
        htmlText += "<tr>";
        for (let c = minCol; c <= maxCol; c++) {
          const td = tr.children[c];
          if (td) {
            let val = td.innerText.trim();
            // Escapar formato TSV
            if (val.includes('\\t') || val.includes('\\n') || val.includes('"')) {
              val = `"${val.replace(/"/g, '""')}"`;
            }
            rowText.push(val);
            htmlText += `<td>${td.innerHTML}</td>`;
          }
        }
        tsvText += rowText.join('\t') + '\n';
        htmlText += "</tr>";
      }
      htmlText += "</tbody></table>";

      if (e.clipboardData) {
        e.clipboardData.setData('text/plain', tsvText);
        e.clipboardData.setData('text/html', htmlText);
      }

      if (window.ClipboardUtil) {
        ClipboardUtil.showToast('Celdas seleccionadas copiadas', 'success', 2000);
      }
    }
  });

})();
