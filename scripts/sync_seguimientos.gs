/**
 * UST 2026-1 — Sincronización automática: DI Seguimientos → Dashboard
 *
 * INSTRUCCIONES:
 * 1. Abre el Dashboard: https://docs.google.com/spreadsheets/d/1NBGFr8UhU5trHSuS87ie1rE9U3rKjapTAQ3oKUq3xn4
 * 2. Extensiones → Apps Script → pega este script → Guarda
 * 3. Ejecutar → sincronizarDashboard → autoriza el acceso
 * 4. Para actualización automática: ícono reloj → Añadir activador → sincronizarDashboard → cada hora
 *
 * PUBLICAR PARA VERCEL:
 * 1. Archivo → Compartir → Publicar en la web
 * 2. Hoja: "DASHBOARD" → formato: Valores separados por comas (.csv) → Publicar
 * 3. Copia la URL generada — la necesitarás para el dashboard Vercel
 */

// ─── ID del Dashboard (fuente de verdad) ─────────────────────────────────────
const DASHBOARD_ID = '1NBGFr8UhU5trHSuS87ie1rE9U3rKjapTAQ3oKUq3xn4';

// ─── IDs de planillas de seguimiento por DI ───────────────────────────────────
const SEGUIMIENTO_IDS = [
  { id: '1mvc5S3viFNOQr3e8oXHKPT5gQHiA0rS5AsXObR83IYA', di: 'Andreina'    },
  { id: '1MzcRjzWPa6PDCkZdTdOGpyMeUJTXtUGMNg3tkIORWiE', di: 'Cristián'    },
  { id: '1-olfhgLaRjo0rT8P9wyCiJ_yPD98RWoL3-bD6J7ofTs', di: 'Brenda'      },
  { id: '1bd8AhhYo0MEuakMQQWeFmyYeWDKTz6FjY-X5BBNDFQ8', di: 'Noelia DI'   },
  { id: '1pZA-Q3WA1ddsNruZQS94fEhncKZUl9ZAJlDSTDZBbls', di: 'Brenda JL'   },
  { id: '1cWGEka3pfO1mJD64IsP-hwbh5bCIy44H8Y9YgalA6jg', di: 'José Luis 1' },
  { id: '133qfGdFmRzBxuTOZ6pWB1T3nTli2C1Ugg2xJtKFRrfM', di: 'José Luis 2' },
  { id: '1rnHRTp8r_emmiOxYffbnUK8xMs-gkJZCWVYGN0Qhd_Q', di: 'José Luis 3' },
  { id: '1RIQ5zJ3tv6qiUJQYeBeDu_-fG9BUmdxxog68GzAku10', di: 'Natalia 1'   },
  { id: '1L43dz5b1_QbAbc1pJSmAzHYoZUGQCmAsksVmhsA-FPc', di: 'Natalia 2'   },
];

// Clasificación de estados → columna del Dashboard
// Columnas Dashboard (0-indexado): JP=0 DI=1 Asignatura=2 Total=3
//   SinComenzar=4 EnProdDI=5 ValidadoDocente=6 EnProdDM=7 QA=8 Terminado=9
//   %Avance=10 BarraProgreso=11
const CLASIFICAR_ESTADO = (estado) => {
  const e = (estado || '').toLowerCase().trim();
  if (e === '' || e === 'sin comenzar')                          return 'sinComenzar';
  if (e.includes('en producción di') || e === 'borrador')       return 'enProdDI';
  if (e.includes('validado') || e === 'terminado')              return 'validadoDocente';
  if (e.includes('en producción dm'))                           return 'enProdDM';
  if (e.includes('implementación') || e.includes('implementacion') || e.includes('en implementación')) return 'qa';
  if (e === 'qa' || e === 'completado')                         return 'terminado';
  return 'sinComenzar';
};

// ─── Función principal ─────────────────────────────────────────────────────────
function sincronizarDashboard() {
  const ui = SpreadsheetApp.getUi();

  // 1. Leer todos los seguimientos DI y construir mapa: asignatura → conteos
  const mapaAsig = {}; // key: normalizado → { sinComenzar, enProdDI, validadoDocente, enProdDM, qa, terminado }

  SEGUIMIENTO_IDS.forEach(entry => {
    try {
      const ss = SpreadsheetApp.openById(entry.id);
      ss.getSheets().forEach(sheet => {
        const resultado = contarRecursosHoja(sheet);
        if (!resultado) return;

        const key = normalizar(resultado.asignatura);
        if (!mapaAsig[key]) {
          mapaAsig[key] = { nombre: resultado.asignatura, ...resultado.conteos };
        } else {
          // Merge: sumar conteos de ambas entradas (si el DI tiene la asig en 2 archivos)
          Object.keys(resultado.conteos).forEach(k => {
            mapaAsig[key][k] = (mapaAsig[key][k] || 0) + resultado.conteos[k];
          });
        }
        Logger.log('✓ ' + resultado.asignatura + ' (' + entry.di + ')');
      });
    } catch (e) {
      Logger.log('✗ ' + entry.id + ' [' + entry.di + ']: ' + e.message);
    }
  });

  // 2. Abrir Dashboard y actualizar filas
  const dashboard = SpreadsheetApp.openById(DASHBOARD_ID);
  const sheet = dashboard.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  // Encontrar fila de encabezados (tiene "JP", "DI", "Asignatura")
  let headerRow = -1;
  for (let i = 0; i < data.length; i++) {
    const fila = data[i].join(' ').toLowerCase();
    if (fila.includes('asignatura') && (fila.includes(' jp') || fila.includes('jp '))) {
      headerRow = i;
      break;
    }
  }

  if (headerRow === -1) {
    ui.alert('No se encontró la fila de encabezados en el Dashboard.');
    return;
  }

  let actualizadas = 0;
  let noEncontradas = [];

  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const asignaturaRaw = String(row[2] || '').trim(); // Columna C = Asignatura
    if (!asignaturaRaw) continue;

    const key = normalizar(asignaturaRaw);
    let datos = mapaAsig[key] || buscarCoinidenciaParcial(mapaAsig, key);

    if (!datos) {
      noEncontradas.push(asignaturaRaw);
      continue;
    }

    const { sinComenzar, enProdDI, validadoDocente, enProdDM, qa, terminado } = datos;
    const total = sinComenzar + enProdDI + validadoDocente + enProdDM + qa + terminado;
    const avancePct = total > 0 ? Math.round(((terminado + qa) / total) * 100) : 0;
    const barra = generarBarra(avancePct);

    // Escribir columnas D a L (índices 3..10)
    const fila = i + 1;
    sheet.getRange(fila, 4).setValue(total);           // D: Total Recursos
    sheet.getRange(fila, 5).setValue(sinComenzar);     // E: Sin Comenzar
    sheet.getRange(fila, 6).setValue(enProdDI);        // F: En Prod. DI
    sheet.getRange(fila, 7).setValue(validadoDocente); // G: Validado Docente
    sheet.getRange(fila, 8).setValue(enProdDM);        // H: En Prod. DM
    sheet.getRange(fila, 9).setValue(qa);              // I: QA/Implementación
    sheet.getRange(fila, 10).setValue(terminado);      // J: Terminado
    sheet.getRange(fila, 11).setValue(avancePct + '%');// K: % Avance
    sheet.getRange(fila, 12).setValue(barra);          // L: Barra

    actualizadas++;
  }

  // Actualizar timestamp en fila 1
  const encabezadoActual = String(sheet.getRange(1, 1).getValue());
  const sinFecha = encabezadoActual.replace(/·\s*Actualizado.*?(?=·|$)/g, '');
  sheet.getRange(1, 1).setValue(sinFecha.trim() + ' · Actualizado ' + Utilities.formatDate(new Date(), 'America/Santiago', 'dd MMM yyyy'));

  const msg = `Dashboard actualizado.\n✓ ${actualizadas} asignaturas\n✗ Sin coincidencia: ${noEncontradas.length}${noEncontradas.length > 0 ? '\n' + noEncontradas.join('\n') : ''}`;
  Logger.log(msg);
  ui.alert(msg);
}

// ─── Contar recursos por estado en una hoja DI ───────────────────────────────
function contarRecursosHoja(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 5) return null;

  // Buscar nombre de asignatura
  let asignatura = null;
  for (let i = 0; i < Math.min(8, data.length); i++) {
    for (let j = 0; j < data[i].length; j++) {
      const cell = String(data[i][j] || '');
      if (cell.match(/^asignatura\s*:/i)) {
        asignatura = cell.replace(/^asignatura\s*:\s*/i, '').trim();
        if (asignatura.includes(',') && asignatura.length > 50) {
          asignatura = asignatura.split(',')[0].trim();
        }
        break;
      }
    }
    if (asignatura) break;
  }
  if (!asignatura) return null;

  // Encontrar columna "Estado"
  let headerIdx = -1, estadoCol = -1, semanaCol = -1;
  for (let i = 0; i < Math.min(15, data.length); i++) {
    let hasSemana = false, hasEstado = false;
    for (let j = 0; j < data[i].length; j++) {
      const h = String(data[i][j] || '').toLowerCase().trim();
      if (h === 'semana') { semanaCol = j; hasSemana = true; }
      if (h === 'estado') { estadoCol = j; hasEstado = true; }
    }
    if (hasSemana && hasEstado) { headerIdx = i; break; }
  }
  if (headerIdx === -1 || estadoCol === -1) return null;

  // Contar por categoría
  const conteos = { sinComenzar: 0, enProdDI: 0, validadoDocente: 0, enProdDM: 0, qa: 0, terminado: 0 };

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    // Solo filas con semana válida (S0, S1, S2, ...)
    const semana = String(row[semanaCol] || '').trim();
    if (!semana.match(/^S\d+$/i)) continue;

    const estado = String(row[estadoCol] || '').trim();
    if (!estado) continue;

    const cat = CLASIFICAR_ESTADO(estado);
    conteos[cat]++;
  }

  return { asignatura, conteos };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function generarBarra(pct) {
  const llenos = Math.round(pct / 5);
  const vacios = 20 - llenos;
  return '█'.repeat(llenos) + '░'.repeat(vacios) + '  ' + pct + '%';
}

function normalizar(nombre) {
  return String(nombre)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function buscarCoinidenciaParcial(mapa, key) {
  const palabras = key.split(' ').slice(0, 4).join(' ');
  for (const [k, v] of Object.entries(mapa)) {
    if (k.includes(palabras) || palabras.includes(k.split(' ').slice(0, 4).join(' '))) return v;
  }
  const palabras3 = key.split(' ').slice(0, 3).join(' ');
  for (const [k, v] of Object.entries(mapa)) {
    if (k.startsWith(palabras3)) return v;
  }
  return null;
}
