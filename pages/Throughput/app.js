document.addEventListener('DOMContentLoaded', () => {
  // Inicialización de autenticación
  if (window.GoogleSheetsService) {
    window.GoogleSheetsService.initAuth();
    updateAuthUI();
  }

  // Elementos UI
  const btnGoogleLogin = document.getElementById('btnGoogleLogin');
  const btnGoogleLogout = document.getElementById('btnGoogleLogout');
  const btnFetchData = document.getElementById('btnFetchData');
  const sheetUrlInput = document.getElementById('sheetUrlInput');
  const loadingIndicator = document.getElementById('loadingIndicator');
  const dashboardSection = document.getElementById('dashboardSection');
  const tabBtns = document.querySelectorAll('.tab-btn');

  // Listeners de Autenticación
  btnGoogleLogin?.addEventListener('click', () => {
    window.GoogleSheetsService.requestAccessToken((success) => {
      if (success) updateAuthUI();
    });
  });

  btnGoogleLogout?.addEventListener('click', () => {
    window.GoogleSheetsService.logout();
    updateAuthUI();
  });

  function updateAuthUI() {
    const isAuth = window.GoogleSheetsService.isAuthenticated();
    document.getElementById('btnGoogleLogin').classList.toggle('hidden', isAuth);
    document.getElementById('userInfo').classList.toggle('hidden', !isAuth);
    if (isAuth && window.GoogleSheetsService.userEmail) {
      document.getElementById('userEmail').textContent = window.GoogleSheetsService.userEmail;
    }
  }

  // Navegación por Pestañas
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.target).classList.add('active');
    });
  });

  // Carga de Datos
  btnFetchData?.addEventListener('click', async () => {
    const urlOrId = sheetUrlInput.value.trim();
    if (!urlOrId) {
      alert("Por favor, ingresa la URL del Google Sheet.");
      return;
    }
    
    const sheetId = window.GoogleSheetsService.extractSpreadsheetId(urlOrId);
    if (!sheetId) {
      alert("No se pudo extraer el ID del documento. Revisa la URL.");
      return;
    }

    if (!window.GoogleSheetsService.isAuthenticated()) {
      alert("Debes acceder con Google primero para leer el archivo.");
      return;
    }

    loadingIndicator.classList.remove('hidden');
    dashboardSection.classList.add('hidden');

    try {
      // Obtener rango Frescos
      const dataFrescos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!A5:L109");
      // Obtener rango Secos
      const dataSecos = await window.GoogleSheetsService.fetchSheetData(sheetId, "BD_Grafico!R5:AA109");

      renderTable('tableFrescos', dataFrescos);
      renderTable('tableSecos', dataSecos);

      // Intentar generar gráficos (asume Primera Col = Etiqueta, resto = series temporales)
      generateCharts('chartFrescos', dataFrescos);
      generateCharts('chartSecos', dataSecos);

      dashboardSection.classList.remove('hidden');
    } catch (error) {
      console.error(error);
      alert("Error al cargar los datos: " + error.message);
    } finally {
      loadingIndicator.classList.add('hidden');
    }
  });

  function renderTable(tableId, data) {
    const thead = document.getElementById(`${tableId}Header`);
    const tbody = document.getElementById(`${tableId}Body`);
    
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (!data || !data.headers || data.headers.length === 0) return;

    // Render Headers
    data.headers.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      thead.appendChild(th);
    });

    // Render Rows
    data.rows.forEach(row => {
      const tr = document.createElement('tr');
      data.headers.forEach(h => {
        const td = document.createElement('td');
        const val = row[h];
        // Formatear números si es posible
        if (typeof val === 'number') {
          td.textContent = val.toLocaleString('es-PE', { maximumFractionDigits: 2 });
        } else {
          td.textContent = val !== null && val !== undefined ? val : '';
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  // Objetos para destruir gráficos previos
  const charts = {};

  function generateCharts(prefix, data) {
    if (!data || !data.rows || data.rows.length === 0) return;

    // Asumimos que la primera columna son las Divisiones y las demás son Semanas/Meses
    const headers = data.headers;
    const labels = data.rows.map(r => r[headers[0]] || 'N/A');
    
    // Para simplificar, si hay 3 métricas en el archivo (Recibo, Despacho, Inventario)
    // vamos a graficar la primera columna de datos numérica en Recibo, etc.
    // Como el usuario indicó que las columnas son Semanas, crearemos un dataset por cada Semana.
    
    // Generamos un gráfico base para Recibo, Despacho, Inventario
    // (Esta es una aproximación genérica para visualizar los datos leídos)
    const datasets = [];
    const colors = [
      'rgba(54, 162, 235, 0.7)', 'rgba(255, 99, 132, 0.7)', 'rgba(75, 192, 192, 0.7)',
      'rgba(255, 206, 86, 0.7)', 'rgba(153, 102, 255, 0.7)', 'rgba(255, 159, 64, 0.7)'
    ];

    for (let i = 1; i < headers.length; i++) {
      const colName = headers[i];
      datasets.push({
        label: colName,
        data: data.rows.map(r => {
          const val = r[colName];
          return typeof val === 'number' ? val : parseFloat(val) || 0;
        }),
        backgroundColor: colors[i % colors.length],
        borderColor: colors[i % colors.length].replace('0.7', '1'),
        borderWidth: 1
      });
    }

    createChart(`${prefix}Recibo`, 'Bar', labels, datasets);
    createChart(`${prefix}Despacho`, 'Line', labels, datasets);
    createChart(`${prefix}Inventario`, 'Bar', labels, datasets);
  }

  function createChart(canvasId, type, labels, datasets) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    if (charts[canvasId]) {
      charts[canvasId].destroy();
    }

    charts[canvasId] = new Chart(ctx, {
      type: type.toLowerCase(),
      data: {
        labels: labels,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top' }
        },
        scales: {
          y: { beginAtZero: true }
        }
      }
    });
  }
});
