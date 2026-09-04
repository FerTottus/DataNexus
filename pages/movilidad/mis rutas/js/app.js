// Archivo: app.js

// 1. Inicializar el mapa
const map = L.map('mapa').setView([-12.046374, -77.042793], 11);

L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© OpenStreetMap contributors'
}).addTo(map);

// 2. Colores (Corregidos 6C y 5B)
const coloresRutas = {
    '1': '#E6194B', '2': '#3CB44B', '3': '#FFE119', '4A': '#4363D8',
    '4B': '#F58231', '5A': '#911EB4', '5C': '#42D4F4', '6A': '#F032E6',
    '6B': '#BFEF45', '7': '#FABED4', '8': '#469990', '9': '#800000',
    '6C': '#9A6324', // Antes 10
    '5B': '#000075'  // Antes 11
};

// Variables globales para interactividad
const líneasDeMapa = {};
const tarjetasDePanel = {};
const marcadoresDeMapa = {}; // NUEVO: Para guardar los pines y etiquetas

// --- FUNCIÓN: RESETEAR VISTA (Clic fuera del mapa) ---
function resetearVista() {
    // 1. Restaurar las líneas
    for (const [id, linea] of Object.entries(líneasDeMapa)) {
        linea.setStyle({ weight: 4, color: coloresRutas[id], opacity: 0.8 });
    }
    // 2. Restaurar los paneles
    document.querySelectorAll('.tarjeta-resaltada').forEach(t => t.classList.remove('tarjeta-resaltada'));
    // 3. Restaurar la opacidad de los pines y sus nombres
    for (const [id, marcadores] of Object.entries(marcadoresDeMapa)) {
        marcadores.forEach(m => {
            m.setOpacity(1);
            const tooltipEl = m.getTooltip()?.getElement();
            if (tooltipEl) tooltipEl.style.opacity = '1';
        });
    }
}

// Detectar clic en el fondo para desmarcar
map.on('click', function (e) {
    if (e.originalEvent.target.id === 'mapa' || e.originalEvent.target.classList.contains('leaflet-container')) {
        resetearVista();
    }
});

// --- FUNCIÓN DE RESALTADO ---
function resaltarRutaSeleccionada(idRutaSeleccionada) {
    // 1. Líneas
    for (const [idRuta, linea] of Object.entries(líneasDeMapa)) {
        const colorOriginal = coloresRutas[idRuta] || '#333333';
        if (idRuta === idRutaSeleccionada) {
            linea.setStyle({ weight: 8, color: colorOriginal, opacity: 1 });
            linea.bringToFront();
        } else {
            linea.setStyle({ weight: 2, color: colorOriginal, opacity: 0.15 });
        }
    }
    // 2. Pines y Etiquetas (Tooltips)
    for (const [idRuta, marcadores] of Object.entries(marcadoresDeMapa)) {
        if (idRuta === idRutaSeleccionada) {
            marcadores.forEach(m => {
                m.setOpacity(1); // 100% visible
                const tooltipEl = m.getTooltip()?.getElement();
                if (tooltipEl) tooltipEl.style.opacity = '1';
                m.setZIndexOffset(1000); // Traer el pin al frente
            });
        } else {
            marcadores.forEach(m => {
                m.setOpacity(0.2); // Pin transparente
                const tooltipEl = m.getTooltip()?.getElement();
                if (tooltipEl) tooltipEl.style.opacity = '0.2'; // Nombre transparente
                m.setZIndexOffset(0);
            });
        }
    }
    // 3. Tarjetas del panel
    document.querySelectorAll('.tarjeta-resaltada').forEach(t => t.classList.remove('tarjeta-resaltada'));
    if (tarjetasDePanel[idRutaSeleccionada]) {
        tarjetasDePanel[idRutaSeleccionada].classList.add('tarjeta-resaltada');
        tarjetasDePanel[idRutaSeleccionada].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
}

// 3. Crear pines SVG
const crearIconoDePinColoreado = (color) => {
    const svgPin = `<svg viewBox="0 0 24 24" fill="${color}" width="24" height="24" xmlns="http://www.w3.org/2000/svg" style="filter: drop-shadow(1px 2px 1px rgba(0,0,0,0.4));"><path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5a2.5 2.5 0 110-5 2.5 2.5 0 010 5z"/></svg>`;
    return L.divIcon({
        html: svgPin, className: 'pin-svg-personalizado', iconSize: [24, 24], iconAnchor: [12, 24]
    });
};

// 4. Agrupar paraderos y dibujar pines
const rutasAgrupadas = {};

paraderos.forEach(punto => {
    if (!rutasAgrupadas[punto.ruta]) {
        rutasAgrupadas[punto.ruta] = [];
        marcadoresDeMapa[punto.ruta] = []; // Inicializar array de pines
    }
    rutasAgrupadas[punto.ruta].push(`${punto.lng},${punto.lat}`);

    const colorDeEstaRuta = coloresRutas[punto.ruta] || '#333333';

    const marcador = L.marker([punto.lat, punto.lng], { icon: crearIconoDePinColoreado(colorDeEstaRuta) })
        .addTo(map)
        .bindTooltip(punto.nombre, {
            permanent: true,
            direction: 'top',
            className: 'etiqueta-paradero',
            offset: [0, -20]
        })
        .on('click', (e) => {
            L.DomEvent.stopPropagation(e);
            resaltarRutaSeleccionada(punto.ruta);
        });

    marcadoresDeMapa[punto.ruta].push(marcador); // Guardamos para opacarlos luego
});

// 5. Solicitar trazado a Mapbox
const tokenMapbox = ['pk', 'eyJ1IjoiZmh1cnRhZG9hIiwiYSI6ImNtbnRmeW52NTBwb2sycW9uYWJjeXd6Mm8ifQ', 'LcHL2SI6zsJ-oQyg3JUFrw'].join('.');
const contenedorLista = document.getElementById('lista-rutas');
contenedorLista.innerHTML = '';

// --- FUNCIÓN CORREGIDA PARA GENERAR LINK DE GOOGLE MAPS ---
function generarLinkGoogle(coordsArray) {
    if (coordsArray.length < 2) return "";
    const origin = coordsArray[0];
    const destination = coordsArray[coordsArray.length - 1];
    const waypoints = coordsArray.slice(1, -1).join('|');

    // URL OFICIAL DE GOOGLE MAPS DIRECTIONS API
    let url = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${destination}`;
    if (waypoints) {
        url += `&waypoints=${encodeURIComponent(waypoints)}`;
    }
    return url;
}

for (const [nombreRuta, coordenadas] of Object.entries(rutasAgrupadas)) {
    if (coordenadas.length < 2) continue;

    const stringCoordenadas = coordenadas.join(';');
    const urlMapbox = `https://api.mapbox.com/directions/v5/mapbox/driving/${stringCoordenadas}?geometries=geojson&access_token=${tokenMapbox}`;

    fetch(urlMapbox).then(r => r.json()).then(data => {
        if (data.code === 'Ok') {
            const rutaData = data.routes[0];
            const colorAsignado = coloresRutas[nombreRuta] || '#333333';
            const distanciaKM = (rutaData.distance / 1000).toFixed(1);
            const tiempoMinutos = Math.round(rutaData.duration / 60);

            const layerRuta = L.geoJSON(rutaData.geometry, {
                style: { color: colorAsignado, weight: 4, opacity: 0.8 }
            }).addTo(map);
            líneasDeMapa[nombreRuta] = layerRuta;

            // División de GPS
            const coordsGoogle = coordenadas.map(c => {
                const [lng, lat] = c.split(','); return `${lat},${lng}`;
            });

            let botonesGPS = "";
            if (coordsGoogle.length <= 10) {
                botonesGPS = `<a href="${generarLinkGoogle(coordsGoogle)}" target="_blank" class="btn-gps">Abrir GPS</a>`;
            } else {
                const parte1 = coordsGoogle.slice(0, 10);
                const parte2 = coordsGoogle.slice(9); // Se repite el punto 9 para conectar
                botonesGPS = `
                    <div style="display:flex; gap:5px;">
                        <a href="${generarLinkGoogle(parte1)}" target="_blank" class="btn-gps" style="flex:1;">GPS Parte 1</a>
                        <a href="${generarLinkGoogle(parte2)}" target="_blank" class="btn-gps" style="flex:1; background-color:#34a853;">GPS Parte 2</a>
                    </div>`;
            }

            const tarjetaRuta = document.createElement('div');
            tarjetaRuta.className = 'tarjeta-logistica';

            // --- DISEÑO DE TARJETA MEJORADO (Más separación y sombras) ---
            tarjetaRuta.style.cssText = `
                border-left: 6px solid ${colorAsignado};
                background-color: #f9f9f9;
                margin-bottom: 15px; /* Mayor separación entre rutas */
                padding: 12px;
                border-radius: 6px;
                font-size: 14px;
                cursor: pointer;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            `;

            tarjetaRuta.innerHTML = `
                <div style="font-weight: bold; color: ${colorAsignado}; font-size: 15px;">RUTA ${nombreRuta}</div>
                <div style="display: flex; justify-content: space-between; margin: 8px 0;">
                    <span>🛣️ ${distanciaKM} km</span><span>⏱️ ${tiempoMinutos} min</span>
                </div>
                ${botonesGPS}
            `;
            tarjetasDePanel[nombreRuta] = tarjetaRuta;
            tarjetaRuta.addEventListener('click', (e) => {
                if (e.target.tagName !== 'A') resaltarRutaSeleccionada(nombreRuta);
            });
            contenedorLista.appendChild(tarjetaRuta);
        }
    });
}