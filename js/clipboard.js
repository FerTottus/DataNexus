/**
 * Clipboard & Toast Utilities for SheetPivot
 * Permite copiar celdas individuales, columnas completas y tablas estructuradas (TSV/HTML)
 * directamente compatibles con Excel, Google Sheets, LibreOffice y portapapeles de texto.
 */

const ClipboardUtil = {
  /**
   * Muestra una notificación tipo Toast en pantalla
   * @param {string} message - Mensaje a mostrar
   * @param {'success'|'info'|'danger'} type - Tipo de notificación
   * @param {number} duration - Duración en milisegundos
   */
  showToast(message, type = 'success', duration = 3000) {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let icon = 'fa-circle-check';
    if (type === 'info') icon = 'fa-circle-info';
    if (type === 'danger') icon = 'fa-triangle-exclamation';

    toast.innerHTML = `
      <i class="fa-solid ${icon}"></i>
      <span>${message}</span>
    `;

    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.2s ease';
      setTimeout(() => toast.remove(), 200);
    }, duration);
  },

  /**
   * Copia un valor de celda individual al portapapeles
   * @param {string|number} value 
   */
  async copyCell(value) {
    try {
      const textToCopy = value !== null && value !== undefined ? String(value) : '';
      await navigator.clipboard.writeText(textToCopy);
      this.showToast(`Celda copiada: "${textToCopy.substring(0, 30)}${textToCopy.length > 30 ? '...' : ''}"`, 'info', 2000);
    } catch (err) {
      console.error('Error al copiar celda:', err);
      this.showToast('No se pudo copiar al portapapeles', 'danger');
    }
  },

  /**
   * Copia una columna completa con su encabezado en formato dual (TSV plano + HTML nativo)
   * Compatible con Excel, Google Sheets y texto plano.
   * @param {string} columnName - Nombre del encabezado
   * @param {Array} values - Array de valores de las filas
   */
  async copyColumn(columnName, values) {
    try {
      const escapeTsv = (val) => {
        const s = val !== null && val !== undefined ? String(val) : '';
        if (s.includes('\t') || s.includes('\n') || s.includes('"')) {
          return `"${s.replace(/"/g, '""')}"`;
        }
        return s;
      };

      const escapeHtmlText = (val) => {
        return String(val !== null && val !== undefined ? val : '')
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;');
      };

      const lines = [escapeTsv(columnName), ...values.map(v => escapeTsv(v))];
      const tsvContent = lines.join('\n');

      const htmlContent = `<table><thead><tr><th>${escapeHtmlText(columnName)}</th></tr></thead><tbody>` +
        values.map(v => `<tr><td>${escapeHtmlText(v)}</td></tr>`).join('') +
        `</tbody></table>`;

      if (navigator.clipboard && window.ClipboardItem) {
        const textBlob = new Blob([tsvContent], { type: 'text/plain' });
        const htmlBlob = new Blob([htmlContent], { type: 'text/html' });
        const clipboardItem = new ClipboardItem({
          'text/plain': textBlob,
          'text/html': htmlBlob
        });
        await navigator.clipboard.write([clipboardItem]);
      } else {
        await navigator.clipboard.writeText(tsvContent);
      }

      this.showToast(`Columna "${columnName}" copiada (${values.length} filas)`, 'success');
    } catch (err) {
      console.error('Error al copiar columna:', err);
      try {
        const lines = [columnName, ...values.map(v => (v !== null && v !== undefined ? String(v) : ''))];
        await navigator.clipboard.writeText(lines.join('\n'));
        this.showToast(`Columna "${columnName}" copiada (${values.length} filas)`, 'success');
      } catch (fallbackErr) {
        this.showToast('Error al copiar la columna', 'danger');
      }
    }
  },

  /**
   * Copia una tabla completa en formato dual (TSV plano + HTML nativo)
   * Esto permite que al pegar en Excel o Google Sheets se peguen las celdas y filas de forma nativa.
   * @param {HTMLTableElement} tableElement - Elemento HTML de la tabla
   */
  async copyTable(tableElement) {
    if (!tableElement) return;

    try {
      const escapeTsv = (val) => {
        const s = val !== null && val !== undefined ? String(val) : '';
        if (s.includes('\t') || s.includes('\n') || s.includes('"')) {
          return `"${s.replace(/"/g, '""')}"`;
        }
        return s;
      };

      // 1. Extraer representación TSV (para texto plano y editores)
      const rows = Array.from(tableElement.querySelectorAll('tr'));
      const tsvLines = rows.map(row => {
        const cells = Array.from(row.querySelectorAll('th, td'));
        return cells.map(cell => {
          let text = '';
          if (cell.tagName.toLowerCase() === 'th') {
            const colName = cell.getAttribute('data-col-name');
            if (colName !== null && colName !== '') {
              text = colName;
            } else {
              const span = cell.querySelector('.th-content > span');
              text = span ? span.innerText.trim() : cell.innerText.trim();
            }
          } else {
            text = cell.getAttribute('data-value');
            if (text === null || text === '') {
              const clone = cell.cloneNode(true);
              clone.querySelectorAll('.btn-col-copy, .th-actions').forEach(b => b.remove());
              text = clone.innerText.trim();
            }
          }
          return escapeTsv(text);
        }).join('\t');
      });
      const tsvText = tsvLines.join('\n');

      // 2. Extraer representación HTML limpia (para pegar con formato nativo de celdas)
      const tableClone = tableElement.cloneNode(true);
      tableClone.querySelectorAll('.th-actions, .btn-col-copy').forEach(b => b.remove());
      tableClone.querySelectorAll('th').forEach(th => {
        const colName = th.getAttribute('data-col-name');
        if (colName) th.innerText = colName;
      });
      tableClone.querySelectorAll('td[data-value]').forEach(td => {
        const raw = td.getAttribute('data-value');
        if (raw !== null && raw !== '') {
          td.innerText = raw;
        }
      });
      const htmlText = tableClone.outerHTML;

      // 3. Escribir usando ClipboardItem si el navegador lo soporta
      if (navigator.clipboard && window.ClipboardItem) {
        const textBlob = new Blob([tsvText], { type: 'text/plain' });
        const htmlBlob = new Blob([htmlText], { type: 'text/html' });
        const clipboardItem = new ClipboardItem({
          'text/plain': textBlob,
          'text/html': htmlBlob
        });
        await navigator.clipboard.write([clipboardItem]);
      } else {
        await navigator.clipboard.writeText(tsvText);
      }

      this.showToast('¡Tabla completa copiada! Lista para pegar en Excel o Google Sheets', 'success');
    } catch (err) {
      console.error('Error al copiar tabla:', err);
      // Fallback simple
      try {
        const plainText = tableElement.innerText;
        await navigator.clipboard.writeText(plainText);
        this.showToast('Tabla copiada (modo texto plano)', 'info');
      } catch (fallbackErr) {
        this.showToast('Error al acceder al portapapeles', 'danger');
      }
    }
  }
};

window.ClipboardUtil = ClipboardUtil;
