/**
 * UST 2026-1 — Sincronización automática: DI Seguimientos → Dashboard
 *
 * INSTRUCCIONES:
 * 1. Abre el Dashboard: https://docs.google.com/spreadsheets/d/1NBGFr8UhU5trHSuS87ie1rE9U3rKjapTAQ3oKUq3xn4
 * 2. Extensiones → Apps Script → pega este script → Guarda
 * 3. Ejecutar → sincronizarDashboard → autoriza el acceso
 * 4. Para actualización automática: ícono reloj → Añadir activador → sincronizarDashboard → cada hora
 *
 * ESTADOS DE RECURSOS (5 estados):
 *   Borrador → columna F
 *   Validado → columna G
 *   DM → columna H
 *   Implementación → columna I
 *   Cargado en Aula → columna J
 *
 * PUBLICAR PARA VERCEL:
 * 1. Archivo → Compartir → Publicar en la web
 * 2. Hoja: "DASHBOARD" → formato: Valores separados por comas (.csv) → Publicar
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

// ─── Clasificación de estados → categoría ─────────────────────────────────────
// Columnas Dashboard: F=borrador  G=validado  H=dm  I=implementacion  J=cargadoAula
const CLASIFICAR_ESTADO = (estado) => {
  const e = (estado || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();

  // Cargado en Aula (terminado)
  if (e.includes('cargado en aula') || e.includes('cargado al aula') || e === 'cargado') return 'cargadoAula';

  // Implementación
  if (e.includes('implementacion') || e === 'impl' || e === 'implementado') return 'implementacion';

  // DM (en producción multimedia)
  if (e === 'dm' || e.includes('en produccion dm') || e.includes('en prod dm') || e.includes('produccion dm')) return 'dm';

  // Validado (aprobado por docente)
  if (e === 'validado' || e === 'aprobado' || e.includes('validado docente')) return 'validado';

  // Borrador (en elaboración por DI, enviado a revisión, ajustes)
  if (e === 'borrador' || e.includes('en produccion di') || e.includes('en prod di') ||
      e.includes('ajuste') || e.includes('en elaboracion') || e.includes('revision')) return 'borrador';

  // Sin comenzar
  return 'sinComenzar';
};

// ─── Función principal ─────────────────────────────────────────────────────────
function sincronizarDashboard() {
  const ui = SpreadsheetApp.getUi();

  // 1. Leer todos los seguimientos DI y construir mapa: asignatura → conteos
  const mapaAsig = {};

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

  // 2. Abrir Dashboard y actualizar encabezados + filas
  const dashboard = SpreadsheetApp.openById(DASHBOARD_ID);
  const sheet = dashboard.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  // Encontrar fila de encabezados
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

  // Actualizar encabezados de columnas E-L con los nuevos nombres
  const filaEncabezado = headerRow + 1;
  sheet.getRange(filaEncabezado, 5).setValue('Sin Comenzar');
  sheet.getRange(filaEncabezado, 6).setValue('Borrador');
  sheet.getRange(filaEncabezado, 7).setValue('Validado');
  sheet.getRange(filaEncabezado, 8).setValue('DM');
  sheet.getRange(filaEncabezado, 9).setValue('Implementación');
  sheet.getRange(filaEncabezado, 10).setValue('Cargado en Aula');
  sheet.getRange(filaEncabezado, 11).setValue('% Avance');
  sheet.getRange(filaEncabezado, 12).setValue('Barra');

  let actualizadas = 0;
  let noEncontradas = [];

  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const asignaturaRaw = String(row[2] || '').trim();
    if (!asignaturaRaw) continue;

    const key = normalizar(asignaturaRaw);
    let datos = mapaAsig[key] || buscarCoincidenciaParcial(mapaAsig, key);

    if (!datos) {
      noEncontradas.push(asignaturaRaw);
      continue;
    }

    const { sinComenzar, borrador, validado, dm, implementacion, cargadoAula } = datos;
    const total = sinComenzar + borrador + validado + dm + implementacion + cargadoAula;

    // %Avance ponderado por etapa del recurso
    const avancePct = total > 0 ? Math.round(
      (cargadoAula * 1.0 + implementacion * 0.75 + dm * 0.5 + validado * 0.25 + borrador * 0.05) / total * 100
    ) : 0;

    const barra = generarBarra(avancePct);

    const fila = i + 1;
    sheet.getRange(fila, 4).setValue(total);
    sheet.getRange(fila, 5).setValue(sinComenzar);
    sheet.getRange(fila, 6).setValue(borrador);
    sheet.getRange(fila, 7).setValue(validado);
    sheet.getRange(fila, 8).setValue(dm);
    sheet.getRange(fila, 9).setValue(implementacion);
    sheet.getRange(fila, 10).setValue(cargadoAula);
    sheet.getRange(fila, 11).setValue(avancePct + '%');
    sheet.getRange(fila, 12).setValue(barra);

    actualizadas++;
  }

  // Actualizar timestamp
  const encabezadoActual = String(sheet.getRange(1, 1).getValue());
  const sinFecha = encabezadoActual.replace(/·\s*Actualizado.*?(?=·|$)/g, '');
  sheet.getRange(1, 1).setValue(sinFecha.trim() + ' · Actualizado ' + Utilities.formatDate(new Date(), 'America/Santiago', 'dd MMM yyyy HH:mm'));

  const msg = `Dashboard actualizado.\n✓ ${actualizadas} asignaturas\n✗ Sin coincidencia: ${noEncontradas.length}${noEncontradas.length > 0 ? '\n' + noEncontradas.join('\n') : ''}`;
  Logger.log(msg);
  ui.alert(msg);
}

// ─── Contar recursos por estado en una hoja DI ───────────────────────────────
// El estado se infiere desde los links/fechas completados, no del campo "Estado"
function contarRecursosHoja(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 5) return null;

  // Buscar nombre de asignatura (celda que empieza con "Asignatura:")
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

  // Encontrar fila de encabezados buscando columnas clave
  let headerIdx = -1;
  let cols = { semana: -1, linkBorrador: -1, linkValidado: -1, linkDM: -1, linkImpl: -1, cargadoMoodle: -1 };

  for (let i = 0; i < Math.min(15, data.length); i++) {
    const fila = data[i].map(c => String(c || '').toLowerCase().replace(/\s+/g, ' ').trim());
    // La fila de encabezado tiene "semana" e "id recurso"
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

  // Contar por estado inferido desde links completados
  const conteos = { sinComenzar: 0, borrador: 0, validado: 0, dm: 0, implementacion: 0, cargadoAula: 0 };

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    const semana = String(row[cols.semana] || '').trim();
    if (!semana.match(/^S\d+$/i)) continue;

    const tieneImpl     = cols.linkImpl      >= 0 && String(row[cols.linkImpl]      || '').trim().length > 5;
    const tieneDM       = cols.linkDM        >= 0 && String(row[cols.linkDM]        || '').trim().length > 5;
    const tieneValidado = cols.linkValidado  >= 0 && String(row[cols.linkValidado]  || '').trim().length > 5;
    const tieneBorrador = cols.linkBorrador  >= 0 && String(row[cols.linkBorrador]  || '').trim().length > 5;
    const cargado       = cols.cargadoMoodle >= 0 && String(row[cols.cargadoMoodle] || '').trim().toLowerCase() === 'sí';

    if (cargado)          conteos.cargadoAula++;
    else if (tieneImpl)   conteos.implementacion++;
    else if (tieneDM)     conteos.dm++;
    else if (tieneValidado) conteos.validado++;
    else if (tieneBorrador) conteos.borrador++;
    else                  conteos.sinComenzar++;
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

function buscarCoincidenciaParcial(mapa, key) {
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
