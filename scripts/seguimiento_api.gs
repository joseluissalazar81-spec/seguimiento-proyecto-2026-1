/**
 * UST 2026-1 — API de Estados de Seguimiento DI
 *
 * INSTRUCCIONES DE PUBLICACIÓN:
 * 1. Abre cualquier planilla Google → Extensiones → Apps Script
 * 2. Pega este script completo → Guarda (Ctrl+S)
 * 3. Clic en "Implementar" → "Nueva implementación"
 * 4. Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier persona
 * 5. Clic "Implementar" → copia la URL que aparece
 * 6. Pégala en el dashboard cuando te la pida
 *
 * La URL tiene este formato:
 *   https://script.google.com/macros/s/XXXXXXX/exec
 */

// ─── IDs de planillas de seguimiento por DI ──────────────────────────────────
const SEG_IDS = [
  { id: '1mvc5S3viFNOQr3e8oXHKPT5gQHiA0rS5AsXObR83IYA', di: 'Andreina'    },
  { id: '1MzcRjzWPa6PDCkZdTdOGpyMeUJTXtUGMNg3tkIORWiE', di: 'Cristián'    },
  { id: '1-olfhgLaRjo0rT8P9wyCiJ_yPD98RWoL3-bD6J7ofTs', di: 'Brenda'      },
  { id: '1bd8AhhYo0MEuakMQQWeFmyYeWDKTz6FjY-X5BBNDFQ8', di: 'Noelia'      },
  { id: '1pZA-Q3WA1ddsNruZQS94fEhncKZUl9ZAJlDSTDZBbls', di: 'Brenda JL'   },
  { id: '1cWGEka3pfO1mJD64IsP-hwbh5bCIy44H8Y9YgalA6jg', di: 'José Luis 1' },
  { id: '133qfGdFmRzBxuTOZ6pWB1T3nTli2C1Ugg2xJtKFRrfM', di: 'José Luis 2' },
  { id: '1rnHRTp8r_emmiOxYffbnUK8xMs-gkJZCWVYGN0Qhd_Q', di: 'José Luis 3' },
  { id: '1RIQ5zJ3tv6qiUJQYeBeDu_-fG9BUmdxxog68GzAku10', di: 'Natalia 1'   },
  { id: '1L43dz5b1_QbAbc1pJSmAzHYoZUGQCmAsksVmhsA-FPc', di: 'Natalia 2'   },
];

// ─── Punto de entrada de la web app ──────────────────────────────────────────
function doGet(e) {
  try {
    const data = leerTodosLosSeguimientos();
    const json = JSON.stringify({ ok: true, data, ts: Date.now() });
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Lee todas las planillas y devuelve mapa asig_norm → conteos ─────────────
function leerTodosLosSeguimientos() {
  const mapa = {};

  SEG_IDS.forEach(entry => {
    try {
      const ss = SpreadsheetApp.openById(entry.id);
      ss.getSheets().forEach(sheet => {
        const res = parsearHoja(sheet, entry.di);
        if (!res) return;
        const key = norm(res.asignatura);
        if (!mapa[key]) {
          mapa[key] = { asignatura: res.asignatura, di: entry.di, sc: 0, di_: 0, vd: 0, dm: 0, qa: 0, ter: 0 };
        }
        mapa[key].sc  += res.sc;
        mapa[key].di_ += res.di;
        mapa[key].vd  += res.vd;
        mapa[key].dm  += res.dm;
        mapa[key].qa  += res.qa;
        mapa[key].ter += res.ter;
      });
    } catch (e) {
      Logger.log('Error ' + entry.di + ': ' + e.message);
    }
  });

  // Renombrar di_ → di en la salida para compatibilidad con el dashboard
  const salida = {};
  Object.entries(mapa).forEach(([k, v]) => {
    salida[k] = { asignatura: v.asignatura, di: v.di, sc: v.sc, bor: v.di_, vd: v.vd, dm: v.dm, qa: v.qa, ter: v.ter };
  });
  return salida;
}

// ─── Parsea una hoja DI y cuenta recursos por estado ─────────────────────────
function parsearHoja(sheet, diNombre) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 4) return null;

  // Buscar "Asignatura:" en las primeras 8 filas
  let asignatura = null;
  for (let i = 0; i < Math.min(8, data.length); i++) {
    for (let j = 0; j < data[i].length; j++) {
      const c = String(data[i][j] || '');
      if (c.match(/^asignatura\s*:/i)) {
        asignatura = c.replace(/^asignatura\s*:\s*/i, '').trim();
        if (asignatura.includes(',') && asignatura.length > 50)
          asignatura = asignatura.split(',')[0].trim();
        break;
      }
    }
    if (asignatura) break;
  }
  if (!asignatura) return null;

  // Encontrar fila de encabezados
  let hdrIdx = -1;
  const cols = { semana: -1, borrador: -1, validado: -1, dm: -1, impl: -1, cargado: -1 };

  for (let i = 0; i < Math.min(15, data.length); i++) {
    const fila = data[i].map(c => String(c || '').toLowerCase().replace(/\s+/g, ' ').trim());
    if (!fila.some(c => c === 'semana')) continue;

    fila.forEach((h, j) => {
      if (h === 'semana')                                           cols.semana   = j;
      // Borrador: col H — cualquier encabezado que contenga 'borrador'
      if (h.includes('borrador'))                                   cols.borrador = j;
      // Validado: col J — contiene 'validado' pero no 'borrador'
      if (h.includes('validado') && !h.includes('borrador'))        cols.validado = j;
      // DM: col L — contiene 'dm' o 'diseño' cerca de 'producto' o 'link'
      if (h.includes('producto dm') || (h.includes('link') && h.includes(' dm')))
                                                                    cols.dm       = j;
      // Implementación: col N
      if (h.includes('producto impl') || h.includes('implementa'))  cols.impl     = j;
      // Cargado en aula: col Q
      if (h.includes('cargado'))                                    cols.cargado  = j;
    });

    // Fallback por posición si los encabezados no matchearon (H=7, J=9, L=11, N=13, Q=16)
    if (cols.borrador === -1 && fila.length > 7)  cols.borrador = 7;
    if (cols.validado === -1 && fila.length > 9)  cols.validado = 9;
    if (cols.dm       === -1 && fila.length > 11) cols.dm       = 11;
    if (cols.impl     === -1 && fila.length > 13) cols.impl     = 13;
    if (cols.cargado  === -1 && fila.length > 16) cols.cargado  = 16;

    hdrIdx = i;
    break;
  }
  if (hdrIdx === -1 || cols.semana === -1) return null;

  // Contar recursos por estado
  let sc = 0, di = 0, vd = 0, dm = 0, qa = 0, ter = 0;

  for (let i = hdrIdx + 1; i < data.length; i++) {
    const row = data[i];
    const semana = String(row[cols.semana] || '').trim().toUpperCase();
    if (!semana.match(/^S\d+$/)) continue;

    const tieneBor  = cols.borrador >= 0 && String(row[cols.borrador] || '').trim().length > 5;
    const tieneVd   = cols.validado >= 0 && String(row[cols.validado] || '').trim().length > 5;
    const tieneDm   = cols.dm       >= 0 && String(row[cols.dm]       || '').trim().length > 5;
    const tieneImpl = cols.impl     >= 0 && String(row[cols.impl]      || '').trim().length > 5;
    const cargadoV  = cols.cargado  >= 0 ? String(row[cols.cargado] || '').trim().toLowerCase() : '';
    const cargado   = cargadoV === 'sí' || cargadoV === 'si' || cargadoV.includes('observac');

    if (cargado)        ter++;
    else if (tieneImpl) qa++;
    else if (tieneDm)   dm++;
    else if (tieneVd)   vd++;
    else if (tieneBor)  di++;
    else                sc++;
  }

  return { asignatura, sc, di, vd, dm, qa, ter };
}

// ─── Normalizar texto ─────────────────────────────────────────────────────────
function norm(s) {
  return String(s)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
