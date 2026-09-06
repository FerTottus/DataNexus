/**
 * Cell Selection Script for SheetPivot
 * Enables Excel-like drag-to-select, auto-scrolling, keyboard navigation (Arrow keys / Shift+Arrows),
 * and copy for table cells.
 */

(function() {
  let isSelecting = false;
  let startCell = null;
  let currentEndCell = null;
  let selectionBounds = null;

  // Variables para auto-scroll durante arrastre
  let activeWrapper = null;
  let autoScrollRaf = null;
  let lastClientX = 0;
  let lastClientY = 0;
  let scrollSpeedX = 0;
  let scrollSpeedY = 0;

  // Add styles dynamically
  const style = document.createElement('style');
  style.innerHTML = `
    .data-table td.cell-selected {
      background-color: rgba(37, 99, 235, 0.25) !important;
      outline: 1.5px solid rgba(37, 99, 235, 0.85);
      outline-offset: -1px;
    }
    .data-table {
      user-select: none; /* Previene la selección de texto nativa mientras arrastras */
    }
  `;
  document.head.appendChild(style);

  function getCellCoords(cell) {
    if (!cell) return null;
    const tr = cell.parentElement;
    if (!tr) return null;
    const tbody = tr.parentElement;
    if (!tbody) return null;
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

  // Actualiza la celda final según las coordenadas actuales del cursor
  function updateCellUnderCursor(clientX, clientY) {
    if (!activeWrapper || !startCell) return;
    const rect = activeWrapper.getBoundingClientRect();

    // Acotar coordenadas dentro del área de datos de la tabla visible
    // Se suma un margen superior (+38px) para evitar el encabezado sticky
    const clampedX = Math.max(rect.left + 5, Math.min(rect.right - 5, clientX));
    const clampedY = Math.max(rect.top + 38, Math.min(rect.bottom - 8, clientY));

    const el = document.elementFromPoint(clampedX, clampedY);
    if (!el) return;
    const td = el.closest('td');
    if (td && td.closest('.data-table') === startCell.table) {
      const coords = getCellCoords(td);
      if (coords && (!currentEndCell || coords.row !== currentEndCell.row || coords.col !== currentEndCell.col)) {
        currentEndCell = coords;
        renderSelection();
      }
    }
  }

  // Bucle de animación para auto-scroll fluido al arrastrar hacia los bordes
  function startAutoScrollLoop() {
    if (autoScrollRaf) return;
    function loop() {
      if (!isSelecting) {
        autoScrollRaf = null;
        return;
      }
      if (activeWrapper && (scrollSpeedX !== 0 || scrollSpeedY !== 0)) {
        activeWrapper.scrollLeft += scrollSpeedX;
        activeWrapper.scrollTop += scrollSpeedY;
        updateCellUnderCursor(lastClientX, lastClientY);
      }
      autoScrollRaf = requestAnimationFrame(loop);
    }
    autoScrollRaf = requestAnimationFrame(loop);
  }

  function stopAutoScrollLoop() {
    if (autoScrollRaf) {
      cancelAnimationFrame(autoScrollRaf);
      autoScrollRaf = null;
    }
    scrollSpeedX = 0;
    scrollSpeedY = 0;
    activeWrapper = null;
  }

  // Mouse Down: Inicio de selección
  document.addEventListener('mousedown', (e) => {
    const td = e.target.closest('td');
    if (td && td.closest('.data-table')) {
      if (e.button !== 0) return; // Solo clic izquierdo

      const targetTable = td.closest('.data-table');

      // Soporte Shift+Click para selección de rango rápido
      if (e.shiftKey && startCell && startCell.table === targetTable) {
        currentEndCell = getCellCoords(td);
        renderSelection();
        return;
      }

      // Limpiamos selecciones previas en todas las tablas
      document.querySelectorAll('.data-table').forEach(t => clearSelection(t));

      isSelecting = true;
      startCell = getCellCoords(td);
      currentEndCell = startCell;
      activeWrapper = td.closest('.table-responsive-wrapper');
      lastClientX = e.clientX;
      lastClientY = e.clientY;
      scrollSpeedX = 0;
      scrollSpeedY = 0;

      renderSelection();
      startAutoScrollLoop();
    } else if (!e.target.closest('.data-table') && !e.target.closest('.btn-col-copy')) {
      // Clic fuera de las tablas y no en botones de copiado, limpiamos
      document.querySelectorAll('.data-table').forEach(t => clearSelection(t));
      startCell = null;
      currentEndCell = null;
      selectionBounds = null;
      stopAutoScrollLoop();
    }
  });

  // Mouse Move en ventana para capturar arrastres fuera del contenedor
  window.addEventListener('mousemove', (e) => {
    if (!isSelecting || !activeWrapper || !startCell) return;

    lastClientX = e.clientX;
    lastClientY = e.clientY;

    const rect = activeWrapper.getBoundingClientRect();
    const EDGE_MARGIN = 50; // Margen de activación de scroll
    const MAX_SPEED = 24;   // Velocidad máxima de scroll por frame

    // Calcular velocidad horizontal
    if (e.clientX > rect.right - EDGE_MARGIN) {
      const overflow = Math.min(e.clientX - (rect.right - EDGE_MARGIN), 100);
      scrollSpeedX = Math.max(6, (overflow / EDGE_MARGIN) * MAX_SPEED);
    } else if (e.clientX < rect.left + EDGE_MARGIN) {
      const overflow = Math.min((rect.left + EDGE_MARGIN) - e.clientX, 100);
      scrollSpeedX = -Math.max(6, (overflow / EDGE_MARGIN) * MAX_SPEED);
    } else {
      scrollSpeedX = 0;
    }

    // Calcular velocidad vertical
    if (e.clientY > rect.bottom - EDGE_MARGIN) {
      const overflow = Math.min(e.clientY - (rect.bottom - EDGE_MARGIN), 100);
      scrollSpeedY = Math.max(6, (overflow / EDGE_MARGIN) * MAX_SPEED);
    } else if (e.clientY < rect.top + EDGE_MARGIN) {
      const overflow = Math.min((rect.top + EDGE_MARGIN) - e.clientY, 100);
      scrollSpeedY = -Math.max(6, (overflow / EDGE_MARGIN) * MAX_SPEED);
    } else {
      scrollSpeedY = 0;
    }

    // Si el cursor está directamente sobre una celda, actualizar al instante
    const td = e.target.closest('td');
    if (td && td.closest('.data-table') === startCell.table) {
      const coords = getCellCoords(td);
      if (coords && (!currentEndCell || coords.row !== currentEndCell.row || coords.col !== currentEndCell.col)) {
        currentEndCell = coords;
        renderSelection();
      }
    } else {
      updateCellUnderCursor(e.clientX, e.clientY);
    }
  });

  // Mouse Up en ventana
  window.addEventListener('mouseup', () => {
    if (isSelecting) {
      isSelecting = false;
      stopAutoScrollLoop();
    }
  });

  // Navegación con teclado (Flechas y Shift+Flechas)
  document.addEventListener('keydown', (e) => {
    // Si el foco está en un campo de texto, select o editable, no intervenir
    const activeEl = document.activeElement;
    const tag = activeEl ? activeEl.tagName.toLowerCase() : '';
    if (tag === 'input' || tag === 'textarea' || tag === 'select' || activeEl?.isContentEditable) {
      return;
    }

    if (!startCell || !startCell.table) return;

    const key = e.key;

    // Ctrl+A o Cmd+A: Seleccionar toda la tabla activa
    if ((e.ctrlKey || e.metaKey) && key.toLowerCase() === 'a') {
      const rows = Array.from(startCell.table.querySelectorAll('tbody tr'));
      if (rows.length > 0) {
        e.preventDefault();
        const maxR = rows.length - 1;
        const maxC = rows[0].children.length - 1;
        startCell = { row: 0, col: 0, table: startCell.table };
        currentEndCell = { row: maxR, col: maxC, table: startCell.table };
        renderSelection();
        return;
      }
    }

    // Escape: Limpiar selección
    if (key === 'Escape') {
      clearSelection(startCell.table);
      startCell = null;
      currentEndCell = null;
      return;
    }

    if (key !== 'ArrowUp' && key !== 'ArrowDown' && key !== 'ArrowLeft' && key !== 'ArrowRight') {
      return;
    }

    // Prevenir el scroll por defecto de la página cuando navegamos celdas
    e.preventDefault();

    const rows = Array.from(startCell.table.querySelectorAll('tbody tr'));
    if (rows.length === 0) return;
    const numRows = rows.length;

    // Tomar la celda de referencia
    const activeRef = currentEndCell || startCell;
    const rowCells = rows[activeRef.row]?.children;
    const maxCols = rowCells ? rowCells.length : (rows[0]?.children.length || 0);
    if (maxCols === 0) return;

    let targetRow = activeRef.row;
    let targetCol = activeRef.col;

    if (key === 'ArrowRight') targetCol = Math.min(maxCols - 1, targetCol + 1);
    if (key === 'ArrowLeft') targetCol = Math.max(0, targetCol - 1);
    if (key === 'ArrowDown') targetRow = Math.min(numRows - 1, targetRow + 1);
    if (key === 'ArrowUp') targetRow = Math.max(0, targetRow - 1);

    if (e.shiftKey) {
      // Expandir o contraer selección manteniendo startCell fijo
      currentEndCell = { row: targetRow, col: targetCol, table: startCell.table };
    } else {
      // Mover selección de celda única
      startCell = { row: targetRow, col: targetCol, table: startCell.table };
      currentEndCell = startCell;
    }

    renderSelection();

    // Desplazar automáticamente para mantener la celda activa en pantalla
    const targetTd = rows[targetRow]?.children[targetCol];
    if (targetTd) {
      targetTd.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  });

  // Interceptar Ctrl+C para copiar celdas seleccionadas (TSV + HTML)
  document.addEventListener('copy', (e) => {
    if (selectionBounds && startCell) {
      e.preventDefault();

      const escapeHtml = (str) => {
        return String(str !== null && str !== undefined ? str : '')
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;');
      };

      const { minRow, maxRow, minCol, maxCol, table } = selectionBounds;
      const rows = table.querySelectorAll('tbody tr');

      let tsvLines = [];
      let htmlText = "<table><tbody>";

      for (let r = minRow; r <= maxRow; r++) {
        const tr = rows[r];
        if (!tr) continue;

        let rowText = [];
        htmlText += "<tr>";
        for (let c = minCol; c <= maxCol; c++) {
          const td = tr.children[c];
          if (td) {
            let val = td.getAttribute('data-value');
            if (val === null || val === '') val = td.innerText.trim();

            let tsvVal = String(val);
            if (tsvVal.includes('\t') || tsvVal.includes('\n') || tsvVal.includes('"')) {
              tsvVal = `"${tsvVal.replace(/"/g, '""')}"`;
            }
            rowText.push(tsvVal);

            htmlText += `<td>${escapeHtml(val)}</td>`;
          }
        }
        tsvLines.push(rowText.join('\t'));
        htmlText += "</tr>";
      }
      htmlText += "</tbody></table>";
      const tsvText = tsvLines.join('\n');

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
