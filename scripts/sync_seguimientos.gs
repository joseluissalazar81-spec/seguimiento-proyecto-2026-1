/**
 * UST 2026-1 — Sincronización automática: DI Seguimientos → Planilla Maestra
 *
 * INSTRUCCIONES:
 * 1. Abre la planilla maestra: https://docs.google.com/spreadsheets/d/14_TBWnenZiLAH0J6Y1dq-UcD3oSA2U2iK-WDtuqfFPY
 * 2. Ve a Extensiones → Apps Script
 * 3. Pega este script completo (reemplaza todo el contenido)
 * 4. Guarda (Ctrl+S)
 * 5. Haz clic en "Ejecutar" → sincronizarSeguimientos
 * 6. Autoriza el acceso cuando te lo pida
 * 7. Para que se actualice solo: haz clic en "Activadores" (ícono reloj) → Añadir activador → syncSeguimientos → cada hora
 */

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

// Orden de etapas (mayor = más avanzado)
const ETAPA = {
  'sin comenzar':      0,
  'en producción di':  1,
  'borrador':          1,
  'en producción dm':  2,
  'validado docente':  3,
  'terminado':         3,
  'implementación':    4,
  'en implementación': 4,
  'qa':                5,
  'completado':        5,
};

// Semanas del Gantt → semanas individuales del seguimiento
const SEMANAS_GANTT = {
  'S0':     ['S0'],
  'S1-S2':  ['S1','S2'],
  'S3-S4':  ['S3','S4'],
  'S5-S6':  ['S5','S6'],
  'S7-S8':  ['S7','S8'],
  'S9-S10': ['S9','S10'],
  'S11-S12':['S11','S12'],
  'S13-S14':['S13','S14'],
  'S15-S16':['S15','S16'],
  'S17-S18':['S17','S18'],
};

// Columnas en la planilla maestra (índice 0)
// Fila de encabezado: Facultad(0) Asignatura(1) Docente(2) Fmt(3) JP(4) DI(5) DG(6) Impl(7)
//   Reunión(8) Planif(9) S0(10) S1-S2(11) S3-S4(12) S5-S6(13) S7-S8(14)
//   S9-S10(15) S11-S12(16) S13-S14(17) S15-S16(18) S17-S18(19) %Avance(20) Estado(21)
const COL = {
  ASIGNATURA:  1,
  REUNION:     8,
  PLANIF:      9,
  S0:          10,
  'S1-S2':     11,
  'S3-S4':     12,
  'S5-S6':     13,
  'S7-S8':     14,
  'S9-S10':    15,
  'S11-S12':   16,
  'S13-S14':   17,
  'S15-S16':   18,
  'S17-S18':   19,
  AVANCE:      20,
  ESTADO:      21,
};

// ─── Función principal ─────────────────────────────────────────────────────────
function sincronizarSeguimientos() {
  const ui = SpreadsheetApp.getUi();
  const master = SpreadsheetApp.getActiveSpreadsheet();

  // Buscar la hoja con gid=1889016436
  const masterSheet = getSheetByGid(master, 1889016436) || master.getSheets()[0];

  Logger.log('Usando hoja maestra: ' + masterSheet.getName());

  // 1. Leer todas las planillas de seguimiento y construir mapa asignatura→estado por semana
  const mapaAsig = {}; // key: nombre normalizado → { ganttStates, reunionRealizada, planifEstado }

  SEGUIMIENTO_IDS.forEach(entry => {
    try {
      const ss = SpreadsheetApp.openById(entry.id);
      const sheets = ss.getSheets();

      sheets.forEach(sheet => {
        const info = parsearHojaSeguimiento(sheet);
        if (!info) return;

        const key = normalizar(info.asignatura);
        if (!mapaAsig[key]) {
          mapaAsig[key] = {
            nombre: info.asignatura,
            semanas: info.semanas,
            planifEstado: info.planifEstado,
          };
        } else {
          // Merge semanas si hay duplicados
          Object.entries(info.semanas).forEach(([s, stats]) => {
            if (!mapaAsig[key].semanas[s]) {
              mapaAsig[key].semanas[s] = stats;
            }
          });
        }

        Logger.log('✓ Procesada: ' + info.asignatura + ' (DI: ' + entry.di + ')');
      });
    } catch (e) {
      Logger.log('✗ Error leyendo ' + entry.id + ' [' + entry.di + ']: ' + e.message);
    }
  });

  // 2. Actualizar planilla maestra
  const masterData = masterSheet.getDataRange().getValues();
  let headerRowIdx = encontrarFilaEncabezado(masterData);

  if (headerRowIdx === -1) {
    ui.alert('No se encontró la fila de encabezados en la planilla maestra.');
    return;
  }

  let actualizadas = 0;
  let noEncontradas = [];

  for (let i = headerRowIdx + 1; i < masterData.length; i++) {
    const row = masterData[i];
    const asignaturaRaw = String(row[COL.ASIGNATURA] || '').trim();
    if (!asignaturaRaw || asignaturaRaw === '') continue;

    // Buscar en el mapa (coincidencia exacta, luego parcial)
    const key = normalizar(asignaturaRaw);
    let diData = mapaAsig[key];

    if (!diData) {
      // Intentar coincidencia parcial
      diData = buscarCoinidenciaParcial(mapaAsig, key);
    }

    if (!diData) {
      noEncontradas.push(asignaturaRaw);
      continue;
    }

    // Calcular estado por semana del Gantt
    let semanasDone = 0;
    let semanasConDatos = 0;

    Object.entries(SEMANAS_GANTT).forEach(([ganttKey, semanasIndiv]) => {
      const colIdx = COL[ganttKey];
      if (!colIdx) return;

      const estado = calcularEstadoGantt(diData.semanas, semanasIndiv);
      masterSheet.getRange(i + 1, colIdx + 1).setValue(estado);

      if (estado !== 'No inicia') semanasConDatos++;
      if (estado === 'Validada' || estado === 'Implementado') semanasDone++;
    });

    // % Avance
    const totalSemanas = Object.keys(SEMANAS_GANTT).length;
    const pct = semanasConDatos > 0 ? Math.round((semanasDone / totalSemanas) * 100) + '%' : '0%';
    masterSheet.getRange(i + 1, COL.AVANCE + 1).setValue(pct);

    // Estado planif (si está disponible)
    if (diData.planifEstado && diData.planifEstado !== '') {
      masterSheet.getRange(i + 1, COL.PLANIF + 1).setValue(diData.planifEstado);
    }

    actualizadas++;
  }

  // Timestamp de última actualización en celda A1 (o similar)
  masterSheet.getRange(1, 1).setNote('Última actualización automática: ' + new Date().toLocaleString('es-CL'));

  const msg = `Sincronización completa.\n\n✓ Asignaturas actualizadas: ${actualizadas}\n✗ No encontradas: ${noEncontradas.length}\n${noEncontradas.length > 0 ? '\nSin coincidencia:\n' + noEncontradas.join('\n') : ''}`;
  Logger.log(msg);
  ui.alert(msg);
}

// ─── Parsear una hoja de seguimiento DI ────────────────────────────────────────
function parsearHojaSeguimiento(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 5) return null;

  // Buscar nombre de asignatura en las primeras 8 filas
  let asignatura = null;
  for (let i = 0; i < Math.min(8, data.length); i++) {
    for (let j = 0; j < data[i].length; j++) {
      const cell = String(data[i][j] || '');
      if (cell.match(/^asignatura\s*:/i)) {
        asignatura = cell.replace(/^asignatura\s*:\s*/i, '').trim();
        // Remover notas adicionales después de coma
        if (asignatura.indexOf(',') > 0 && asignatura.length > 50) {
          asignatura = asignatura.split(',')[0].trim();
        }
        break;
      }
    }
    if (asignatura) break;
  }

  if (!asignatura || asignatura === '') return null;

  // Buscar fila de encabezados de columnas (tiene "Semana" y "Estado")
  let headerIdx = -1;
  let semanaCol = -1, estadoCol = -1, moodleCol = -1;

  for (let i = 0; i < Math.min(15, data.length); i++) {
    const row = data[i];
    let foundSemana = false, foundEstado = false;
    for (let j = 0; j < row.length; j++) {
      const h = String(row[j] || '').toLowerCase().trim();
      if (h === 'semana') { semanaCol = j; foundSemana = true; }
      if (h === 'estado') { estadoCol = j; foundEstado = true; }
      if (h.includes('cargado') || (h.includes('moodle') && h.includes('en'))) moodleCol = j;
    }
    if (foundSemana && foundEstado) { headerIdx = i; break; }
  }

  if (headerIdx === -1 || semanaCol === -1 || estadoCol === -1) return null;

  // Leer recursos y agrupar por semana
  const semanas = {}; // { S0: { total, validados, impl, moodle } }
  let planifEstado = '';

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    const semanaRaw = String(row[semanaCol] || '').trim();
    const estadoRaw = String(row[estadoCol] || '').trim().toLowerCase();

    if (!semanaRaw.match(/^S\d+$/i)) continue;
    const semana = semanaRaw.toUpperCase();

    // Estado normalizado
    const etapa = obtenerEtapa(estadoRaw);

    if (!semanas[semana]) {
      semanas[semana] = { total: 0, minEtapa: 99, maxEtapa: 0, moodle: 0 };
    }

    semanas[semana].total++;
    semanas[semana].minEtapa = Math.min(semanas[semana].minEtapa, etapa);
    semanas[semana].maxEtapa = Math.max(semanas[semana].maxEtapa, etapa);

    if (moodleCol >= 0) {
      const m = String(row[moodleCol] || '').toLowerCase().trim();
      if (m === 'sí' || m === 'si' || m === 'yes' || m === 's') semanas[semana].moodle++;
    }

    // Detectar estado de planificación (recurso S0_Plan)
    const idCol = semanaCol + 2;
    if (idCol < row.length) {
      const id = String(row[idCol] || '').toLowerCase();
      if (id.includes('plan') && semana === 'S0') {
        planifEstado = etapaATextoMaestro(etapa);
      }
    }
  }

  return { asignatura, semanas, planifEstado };
}

// ─── Calcular estado para un grupo de semanas del Gantt ────────────────────────
function calcularEstadoGantt(semanas, semanasIndiv) {
  let totalRecursos = 0;
  let minEtapa = 99;
  let maxEtapa = 0;
  let totalMoodle = 0;

  semanasIndiv.forEach(s => {
    const stats = semanas[s];
    if (!stats || stats.total === 0) return;
    totalRecursos += stats.total;
    minEtapa = Math.min(minEtapa, stats.minEtapa);
    maxEtapa = Math.max(maxEtapa, stats.maxEtapa);
    totalMoodle += stats.moodle;
  });

  if (totalRecursos === 0) return 'No inicia';

  // Todos implementados (en Moodle)
  if (totalMoodle >= totalRecursos) return 'Implementado';

  // Todos al menos validados por docente
  if (minEtapa >= ETAPA['validado docente']) return 'Validada';

  // Alguno está en producción o más avanzado
  if (maxEtapa >= ETAPA['en producción di']) return 'En construcción';

  return 'No inicia';
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function obtenerEtapa(estado) {
  const lower = estado.toLowerCase().trim();
  if (ETAPA[lower] !== undefined) return ETAPA[lower];

  // Búsqueda parcial
  for (const [key, val] of Object.entries(ETAPA)) {
    if (lower.includes(key)) return val;
  }
  return 0;
}

function etapaATextoMaestro(etapa) {
  if (etapa >= ETAPA['implementación']) return 'Validada';
  if (etapa >= ETAPA['validado docente']) return 'Validada';
  if (etapa >= ETAPA['en producción di']) return 'En construcción';
  return 'Enviada a validación DEL';
}

function normalizar(nombre) {
  return String(nombre)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // quitar acentos
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function buscarCoinidenciaParcial(mapa, key) {
  // Intentar con primeras 4 palabras
  const palabras = key.split(' ').slice(0, 4).join(' ');
  for (const [k, v] of Object.entries(mapa)) {
    if (k.includes(palabras) || palabras.includes(k.split(' ').slice(0, 4).join(' '))) {
      return v;
    }
  }

  // Intentar con 3 palabras
  const palabras3 = key.split(' ').slice(0, 3).join(' ');
  for (const [k, v] of Object.entries(mapa)) {
    if (k.startsWith(palabras3)) return v;
  }

  return null;
}

function encontrarFilaEncabezado(data) {
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const fila = row.join(' ').toLowerCase();
    if (fila.includes('asignatura') && fila.includes('facultad')) return i;
  }
  return -1;
}

function getSheetByGid(spreadsheet, gid) {
  const sheets = spreadsheet.getSheets();
  return sheets.find(s => s.getSheetId() === gid) || null;
}
