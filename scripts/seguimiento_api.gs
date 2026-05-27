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

// Estados en la col F del seguimiento → categoría interna
const ESTADO_MAP = {
  'sin comenzar':          'sc',
  'en produccion di':      'di',
  'en produccion borrador':'di',
  'borrador':              'di',
  'validado docente':      'vd',
  'validado':              'vd',
  'en produccion dm':      'dm',
  'en diseno grafico':     'dm',
  'implementacion':        'qa',
  'en implementacion':     'qa',
  'terminado':             'ter',
  'cargado en aula':       'ter',
  'cargado en moodle':     'ter',
};

function doGet(e) {
  try {
    const data = leerTodosLosSeguimientos();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data, ts: Date.now() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function leerTodosLosSeguimientos() {
  const mapa = {};
  SEG_IDS.forEach(entry => {
    try {
      const ss = SpreadsheetApp.openById(entry.id);
      ss.getSheets().forEach(sheet => {
        const res = parsearHoja(sheet);
        if (!res) return;
        const key = norm(res.asignatura);
        if (!mapa[key]) mapa[key] = { asignatura: res.asignatura, di: entry.di, sc:0, di_:0, vd:0, dm:0, qa:0, ter:0 };
        mapa[key].sc  += res.sc;
        mapa[key].di_ += res.di;
        mapa[key].vd  += res.vd;
        mapa[key].dm  += res.dm;
        mapa[key].qa  += res.qa;
        mapa[key].ter += res.ter;
      });
    } catch(e) { Logger.log('Error ' + entry.di + ': ' + e.message); }
  });
  const salida = {};
  Object.entries(mapa).forEach(([k,v]) => {
    salida[k] = { asignatura:v.asignatura, di:v.di, sc:v.sc, bor:v.di_, vd:v.vd, dm:v.dm, qa:v.qa, ter:v.ter };
  });
  return salida;
}

function parsearHoja(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 4) return null;

  // Buscar "Asignatura:" en las primeras 8 filas
  let asignatura = null;
  for (let i = 0; i < Math.min(8, data.length); i++) {
    for (let j = 0; j < data[i].length; j++) {
      const c = String(data[i][j] || '');
      if (c.match(/^asignatura\s*:/i)) {
        asignatura = c.replace(/^asignatura\s*:\s*/i, '').trim();
        if (asignatura.includes(',') && asignatura.length > 50) asignatura = asignatura.split(',')[0].trim();
        break;
      }
    }
    if (asignatura) break;
  }
  if (!asignatura) return null;

  // Encontrar fila de encabezados (busca fila con "Semana" y "Estado")
  let hdrIdx = -1, colSemana = -1, colEstado = -1;

  for (let i = 0; i < Math.min(15, data.length); i++) {
    const fila = data[i].map(c => String(c||'').toLowerCase().replace(/\s+/g,' ').trim());
    if (!fila.some(c => c === 'semana')) continue;
    fila.forEach((h, j) => {
      if (h === 'semana') colSemana = j;
      if (h === 'estado' || h.includes('estado del')) colEstado = j;
    });
    if (colEstado === -1) colEstado = 5; // fallback: col F
    hdrIdx = i;
    break;
  }
  if (hdrIdx === -1 || colSemana === -1) return null;

  // Contar por estado
  let sc=0, di=0, vd=0, dm=0, qa=0, ter=0;
  for (let i = hdrIdx + 1; i < data.length; i++) {
    const row = data[i];
    const semana = String(row[colSemana] || '').trim().toUpperCase();
    if (!semana.match(/^S\d+$/)) continue;

    const estadoRaw = norm(String(row[colEstado] || ''));
    const cat = ESTADO_MAP[estadoRaw] || 'sc';

    if      (cat === 'ter') ter++;
    else if (cat === 'qa')  qa++;
    else if (cat === 'dm')  dm++;
    else if (cat === 'vd')  vd++;
    else if (cat === 'di')  di++;
    else                    sc++;
  }
  return { asignatura, sc, di, vd, dm, qa, ter };
}

function norm(s) {
  return String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'')
    .replace(/[^a-z0-9\s]/g,'').replace(/\s+/g,' ').trim();
}

function debugHoja() {
  const ss = SpreadsheetApp.openById('1cWGEka3pfO1mJD64IsP-hwbh5bCIy44H8Y9YgalA6jg');
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  Logger.log('Hoja: ' + sheet.getName() + ' | Filas: ' + data.length);
  for (let i = 0; i < Math.min(20, data.length); i++) {
    const fila = data[i].map(c => String(c||'').trim()).filter(c=>c);
    if (fila.length > 0) Logger.log('F'+(i+1)+': '+fila.slice(0,8).join(' | '));
  }
}
