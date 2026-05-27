/**
 * UST 2026-1 — Vincular Seguimientos DI → Archivo Reuniones
 *
 * INSTRUCCIONES:
 * 1. Abre el archivo Reuniones:
 *    https://docs.google.com/spreadsheets/d/14_TBWnenZiLAH0J6Y1dq-UcD3oSA2U2iK-WDtuqfFPY
 * 2. Extensiones → Apps Script → pega este script → Guarda
 * 3. Ejecutar → actualizarEstadosSemanas → autoriza el acceso
 * 4. Para automatizar: ícono reloj → Añadir activador → actualizarEstadosSemanas → cada hora
 *
 * QUÉ HACE:
 * Lee el seguimiento de cada DI (links en col H/J/L/N y Cargado en Moodle),
 * determina el estado por semana y lo escribe en las columnas K-T del archivo reuniones.
 *
 * ESTADOS (en orden de avance):
 *   No aplica → No inicia → En construcción → Validada → En implementación → Cargado en aula
 *
 * REGLA: si en un bloque de 2 semanas hay varios recursos, se muestra el PEOR estado
 * (el que menos avanzó), para reflejar que el bloque no está completo hasta que todos avancen.
 */

// ─── IDs de seguimiento por DI ────────────────────────────────────────────────
const SEGUIMIENTO_IDS_V = [
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

// ─── Archivo Reuniones ────────────────────────────────────────────────────────
const REUNIONES_SS_ID = '14_TBWnenZiLAH0J6Y1dq-UcD3oSA2U2iK-WDtuqfFPY';
const REUNIONES_GID   = 1889016436;

// ─── Columnas de semanas en el archivo Reuniones (1-indexed) ─────────────────
// Col K(11)=S0 · Col L(12)=S1-S2 · Col M(13)=S3-S4 · ... · Col T(20)=S17-S18
const SEMANA_COLS = [
  { col: 11, semanas: ['S0']          },
  { col: 12, semanas: ['S1',  'S2']   },
  { col: 13, semanas: ['S3',  'S4']   },
  { col: 14, semanas: ['S5',  'S6']   },
  { col: 15, semanas: ['S7',  'S8']   },
  { col: 16, semanas: ['S9',  'S10']  },
  { col: 17, semanas: ['S11', 'S12']  },
  { col: 18, semanas: ['S13', 'S14']  },
  { col: 19, semanas: ['S15', 'S16']  },
  { col: 20, semanas: ['S17', 'S18']  },
];

// ─── Orden de estados ─────────────────────────────────────────────────────────
const ESTADOS_V = [
  'No inicia',
  'En construcción',
  'Validada',
  'En diseño gráfico',
  'En implementación',
  'Cargado en aula',
];
const ESTADO_ORDEN_V = Object.fromEntries(ESTADOS_V.map((e, i) => [e, i]));

// ─── FUNCIÓN PRINCIPAL ────────────────────────────────────────────────────────
function actualizarEstadosSemanas() {
  const ui = SpreadsheetApp.getUi();

  // 1. Leer todos los seguimientos DI y construir mapa asignatura → semanas
  ui.alert('Iniciando sincronización... (puede tardar 1-2 minutos)');
  const mapaAsig = construirMapaAsignaturas();
  Logger.log('Asignaturas encontradas: ' + Object.keys(mapaAsig).length);

  // 2. Abrir hoja de reuniones
  const ss    = SpreadsheetApp.openById(REUNIONES_SS_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === REUNIONES_GID) || ss.getSheets()[0];
  const data  = sheet.getDataRange().getValues();

  // 3. Encontrar fila de encabezados (busca la que tenga "asignatura" en col B)
  let headerRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1] || '').toLowerCase().includes('asignatura')) {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) {
    ui.alert('Error: no se encontró la fila de encabezados (buscando "Asignatura" en col B).');
    return;
  }

  // 4. Actualizar cada fila de datos
  let actualizadas = 0;
  const noEncontradas = [];

  for (let i = headerRow + 1; i < data.length; i++) {
    const asignaturaRaw = String(data[i][1] || '').trim(); // col B
    if (!asignaturaRaw) continue;

    const key      = normalizarV(asignaturaRaw);
    let datosAsig  = mapaAsig[key];

    // Búsqueda parcial si no hay coincidencia exacta
    if (!datosAsig) {
      const palabras = key.split(' ').slice(0, 3).join(' ');
      for (const [k, v] of Object.entries(mapaAsig)) {
        if (k.includes(palabras) || palabras.includes(k.split(' ').slice(0, 3).join(' '))) {
          datosAsig = v;
          break;
        }
      }
    }

    if (!datosAsig) {
      noEncontradas.push(asignaturaRaw);
      continue;
    }

    // Escribir estado en cada columna de semana
    const filaSheet = i + 1;
    SEMANA_COLS.forEach(({ col, semanas }) => {
      const estado = calcularEstadoBloque(datosAsig, semanas);
      sheet.getRange(filaSheet, col).setValue(estado);
    });

    actualizadas++;
  }

  // 5. Aplicar validación dropdown a todas las columnas de semanas
  aplicarDropdown(sheet, headerRow + 2, data.length);

  const resumen = [
    `✓ ${actualizadas} asignaturas actualizadas`,
    noEncontradas.length > 0
      ? `✗ Sin coincidencia (${noEncontradas.length}):\n${noEncontradas.join('\n')}`
      : '✓ Todas las asignaturas encontradas'
  ].join('\n\n');

  Logger.log(resumen);
  ui.alert(resumen);
}

// ─── Construir mapa: asig_norm → { S0: [estado,...], S1: [...], ... } ─────────
function construirMapaAsignaturas() {
  const mapa = {};

  SEGUIMIENTO_IDS_V.forEach(entry => {
    try {
      const ss = SpreadsheetApp.openById(entry.id);
      ss.getSheets().forEach(sheet => {
        const resultado = leerHojaDI(sheet);
        if (!resultado) return;

        const key = normalizarV(resultado.asignatura);
        if (!mapa[key]) mapa[key] = {};

        Object.entries(resultado.semanas).forEach(([sem, estados]) => {
          if (!mapa[key][sem]) mapa[key][sem] = [];
          mapa[key][sem].push(...estados);
        });

        Logger.log('✓ ' + resultado.asignatura + ' (' + entry.di + ')');
      });
    } catch (e) {
      Logger.log('✗ ' + entry.di + ': ' + e.message);
    }
  });

  return mapa;
}

// ─── Leer una hoja DI y agrupar estados por semana ───────────────────────────
function leerHojaDI(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 5) return null;

  // Buscar nombre de asignatura (celda que empiece con "Asignatura:")
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

  // Encontrar fila de encabezados (busca fila con "Semana" e "ID Recurso")
  let headerIdx = -1;
  const cols = { semana: -1, linkBorrador: -1, linkValidado: -1, linkDM: -1, linkImpl: -1, cargadoMoodle: -1 };

  for (let i = 0; i < Math.min(15, data.length); i++) {
    const fila = data[i].map(c => String(c || '').toLowerCase().replace(/\s+/g, ' ').trim());
    if (!fila.some(c => c === 'semana') || !fila.some(c => c.includes('id recurso') || c === 'id')) continue;

    fila.forEach((h, j) => {
      if (h === 'semana')                                         cols.semana        = j;
      if (h.includes('link recurso') && h.includes('borrador'))  cols.linkBorrador  = j;
      if (h.includes('link recurso') && h.includes('validado'))  cols.linkValidado  = j;
      if (h.includes('link') && h.includes('producto dm'))       cols.linkDM        = j;
      if (h.includes('link') && h.includes('producto impl'))     cols.linkImpl      = j;
      if (h.includes('cargado'))                                  cols.cargadoMoodle = j;
    });
    headerIdx = i;
    break;
  }
  if (headerIdx === -1 || cols.semana === -1) return null;

  // Agrupar recursos por semana con su estado
  const semanas = {};

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row    = data[i];
    const semana = String(row[cols.semana] || '').trim().toUpperCase();
    if (!semana.match(/^S\d+$/)) continue;

    const tieneBorrador = cols.linkBorrador  >= 0 && String(row[cols.linkBorrador]  || '').trim().length > 5;
    const tieneValidado = cols.linkValidado  >= 0 && String(row[cols.linkValidado]  || '').trim().length > 5;
    const tieneDM       = cols.linkDM        >= 0 && String(row[cols.linkDM]        || '').trim().length > 5;
    const tieneImpl     = cols.linkImpl      >= 0 && String(row[cols.linkImpl]      || '').trim().length > 5;
    const cargado       = cols.cargadoMoodle >= 0 &&
                          String(row[cols.cargadoMoodle] || '').trim().toLowerCase() === 'sí';

    let estado;
    if (cargado)            estado = 'Cargado en aula';
    else if (tieneImpl)     estado = 'En implementación';
    else if (tieneDM)       estado = 'En diseño gráfico';
    else if (tieneValidado) estado = 'Validada';
    else if (tieneBorrador) estado = 'En construcción';
    else                    estado = 'No inicia';

    if (!semanas[semana]) semanas[semana] = [];
    semanas[semana].push(estado);
  }

  return { asignatura, semanas };
}

// ─── Calcular estado de un bloque (peor estado de todas las semanas del bloque) ──
function calcularEstadoBloque(datosAsig, semanas) {
  let minOrden   = Infinity;
  let hayRecursos = false;

  semanas.forEach(sem => {
    (datosAsig[sem] || []).forEach(estado => {
      hayRecursos = true;
      const orden = ESTADO_ORDEN_V[estado] ?? 0;
      if (orden < minOrden) minOrden = orden;
    });
  });

  if (!hayRecursos) return 'No aplica';
  return ESTADOS_V[minOrden] || 'No inicia';
}

// ─── Aplicar dropdown a columnas K-T ─────────────────────────────────────────
function aplicarDropdown(sheet, firstDataRow, lastRow) {
  const opciones = ['No aplica', 'No inicia', 'En construcción', 'Validada', 'En diseño gráfico', 'En implementación', 'Cargado en aula'];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();

  const numRows = Math.max(1, lastRow - firstDataRow + 1);
  SEMANA_COLS.forEach(({ col }) => {
    sheet.getRange(firstDataRow, col, numRows, 1).setDataValidation(rule);
  });
}

// ─── Helper normalizar ────────────────────────────────────────────────────────
function normalizarV(nombre) {
  return String(nombre)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
