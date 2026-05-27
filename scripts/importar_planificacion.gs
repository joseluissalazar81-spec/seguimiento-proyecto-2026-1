/**
 * importar_planificacion.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * PRUEBA: Seguimiento 2 José Luis — PLA-S0113 Sistema de producción de frutales
 *
 * INSTRUCCIONES:
 * 1. Abre el Google Sheet de seguimiento JL2:
 *    https://docs.google.com/spreadsheets/d/133qfGdFmRzBxuTOZ6pWB1T3nTli2C1Ugg2xJtKFRrfM
 * 2. Extensiones → Apps Script → pega este código → Guarda (Ctrl+S)
 * 3. Ejecuta una vez "crearMenu" para registrar el menú (o recarga la hoja)
 * 4. Aparece menú "📋 Planificación" en la barra → clic → Importar recursos
 * 5. Pega el ID del archivo Excel de planificación IMPL → OK
 *
 * QUÉ HACE:
 * - Lee la hoja "Planificación por unidades" del Excel de planificación
 * - Extrae todos los recursos de la columna H (T1:/T2:/T3:/T4:)
 * - Genera filas con: Unidad, Semana, ID, Nombre, Tipo, Estado=Sin comenzar
 * - Inserta SOLO los recursos que aún NO existen en el seguimiento
 * - Muestra resumen de cuántos se agregaron
 */

// ── Configuración ─────────────────────────────────────────────────────────────

const CONFIG = {
  // ID del seguimiento JL2 (ya estás aquí)
  SEGUIMIENTO_ID: '133qfGdFmRzBxuTOZ6pWB1T3nTli2C1Ugg2xJtKFRrfM',

  // Columnas del seguimiento (índice 0)
  COL_UNIDAD:    0,   // A — Unidad
  COL_SEMANA:    1,   // B — Semana
  COL_ID:        2,   // C — ID Recurso
  COL_NOMBRE:    3,   // D — Nombre del recurso
  COL_TIPO:      4,   // E — Tipo de recurso
  COL_ESTADO:    5,   // F — Estado
  COL_RESP:      6,   // G — Responsable
  TOTAL_COLS:    20,  // Total de columnas en el seguimiento

  // Email del DI para el campo Responsable
  DI_EMAIL: 'joseluissalazar81@gmail.com',

  // Fila a partir de la cual se insertan los datos (después de headers)
  // El script busca la última fila con datos y agrega debajo
};

// ── Mapeo: tipo (T1-T4) + palabra clave → nombre descriptivo del recurso ─────
const TIPO_MAP = {
  'T1': {
    'video':         'Videoclase',
    'capsula':       'Cápsula',
    'podcast':       'Podcast',
    'control':       'Control',
    'cuestionario':  'Cuestionario',
    'infografia':    'Infografía',
    'genially':      'Genially',
    'interactivo':   'Interactivo',
    '_default':      'Audiovisual',
  },
  'T2': {
    'taller':        'Taller',
    'actividad':     'Actividad de clase',
    'tarea':         'Tarea',
    'rubrica':       'Rúbrica',
    'lista de cotejo':'Lista de cotejo',
    'foro':          'Foro',
    'prueba':        'Prueba de desarrollo',
    'apunte':        'Apunte',
    '_default':      'Actividad',
  },
  'T3': {
    'presentacion':  'Presentación',
    'actividad sincr':'Actividad sincrónica',
    '_default':      'Presentación',
  },
  'T4': {
    'link':          'Lectura',
    'video':         'Recurso video',
    'actividad':     'Actividad TPE',
    'guia':          'Guía de preparación',
    'cuestionario':  'Cuestionario diagnóstico',
    'encuesta':      'Encuesta',
    '_default':      'Recurso TPE',
  },
};

// ── Menú ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Planificación')
    .addItem('Importar recursos desde Excel', 'importarDesdeExcel')
    .addItem('Ver resumen de recursos', 'verResumen')
    .addToUi();
}

function crearMenu() {
  onOpen();
}

// ── Función principal ─────────────────────────────────────────────────────────

function importarDesdeExcel() {
  const ui = SpreadsheetApp.getUi();

  // Pedir el ID o URL del archivo de planificación
  const resp = ui.prompt(
    '📥 Importar planificación',
    'Pega el ID o URL completa del Excel de planificación (versión IMPL):\n\n' +
    'Ejemplo ID: 1U1DAZC-QjcmdG0yeJtfq28YKbC5GBsGH\n' +
    'O URL: https://drive.google.com/file/d/XXXXX/view',
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const input = resp.getResponseText().trim();
  if (!input) {
    ui.alert('❌ No ingresaste ningún ID o URL.');
    return;
  }

  // Extraer el ID del archivo
  const fileId = extraerFileId(input);
  if (!fileId) {
    ui.alert('❌ No se pudo extraer el ID del archivo. Verifica el link.');
    return;
  }

  try {
    ui.alert('⏳ Procesando...', 'Leyendo planificación y generando filas. Esto puede tardar unos segundos.', ui.ButtonSet.OK);

    // Convertir Excel a Google Sheets temporalmente
    const tempSheet = convertirExcelASheet(fileId);
    if (!tempSheet) {
      ui.alert('❌ No se pudo abrir el archivo. Verifica que tengas acceso a él.');
      return;
    }

    // Extraer recursos de la planificación
    const recursos = extraerRecursos(tempSheet);

    // Eliminar el sheet temporal
    DriveApp.getFileById(tempSheet.getId()).setTrashed(true);

    if (recursos.length === 0) {
      ui.alert('⚠️ No se encontraron recursos T1/T2/T3/T4 en la columna H de la planificación.');
      return;
    }

    // Insertar en el seguimiento
    const resultado = insertarEnSeguimiento(recursos);

    // Mostrar resumen
    ui.alert(
      '✅ Importación completada',
      `Recursos encontrados en la planificación: ${recursos.length}\n` +
      `Filas agregadas al seguimiento: ${resultado.agregadas}\n` +
      `Ya existían (no duplicadas): ${resultado.duplicadas}\n\n` +
      `Distribución:\n${resultado.resumen}`,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ Error: ' + e.message);
    Logger.log(e);
  }
}

// ── Extraer ID del archivo desde URL o ID directo ─────────────────────────────

function extraerFileId(input) {
  // Si es una URL de Drive
  const match = input.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (match) return match[1];

  // Si es un ID directo (sin /d/)
  if (/^[a-zA-Z0-9_-]{25,}$/.test(input)) return input;

  return null;
}

// ── Convertir Excel a Google Sheet temporal ───────────────────────────────────

function convertirExcelASheet(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();

    // Si ya es Google Sheets, abrirlo directamente
    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
      return SpreadsheetApp.openById(fileId);
    }

    // Si es Excel, convertirlo
    const copy = Drive.Files.copy(
      { name: '_temp_planificacion_import', mimeType: 'application/vnd.google-apps.spreadsheet' },
      fileId
    );
    return SpreadsheetApp.openById(copy.id);

  } catch (e) {
    Logger.log('Error abriendo archivo: ' + e);
    return null;
  }
}

// ── Extraer recursos de la planificación ─────────────────────────────────────

function extraerRecursos(planSS) {
  // Buscar la hoja "Planificación por unidades"
  let hoja = planSS.getSheetByName('Planificación por unidades');
  if (!hoja) {
    // Intentar nombres alternativos
    const hojas = planSS.getSheets().map(h => h.getName());
    Logger.log('Hojas disponibles: ' + hojas.join(', '));
    hoja = planSS.getSheets().find(h =>
      h.getName().toLowerCase().includes('planificaci') ||
      h.getName().toLowerCase().includes('unidad')
    );
  }
  if (!hoja) throw new Error('No se encontró la hoja "Planificación por unidades"');

  const datos = hoja.getDataRange().getValues();
  const recursos = [];

  let unidadActual = '';
  let semanaActual = '';
  let contadorPorSemana = {}; // {U1_S1: 1, U1_S2: 2, ...}

  for (let i = 3; i < datos.length; i++) { // desde fila 4 (índice 3)
    const fila = datos[i];

    // Columnas (0-based): A=0 RA, B=1 Unidad, C=2 Semana, F=5 Momento, H=7 Recursos
    const unidadRaw  = String(fila[1] || '').trim();
    const semanaRaw  = String(fila[2] || '').trim();
    const momento    = String(fila[5] || '').trim();
    const recursosH  = String(fila[7] || '').trim();

    if (unidadRaw) unidadActual = unidadRaw;
    if (semanaRaw) semanaActual = semanaRaw;

    if (!recursosH || !unidadActual || !semanaActual) continue;

    // Normalizar unidad y semana
    const unidadNorm = normalizarUnidad(unidadActual);
    const semanaNorm = normalizarSemana(semanaActual);
    const claveUS = `${unidadNorm}_${semanaNorm}`;

    // Parsear cada línea de la celda H buscando T1:/T2:/T3:/T4:
    const lineas = recursosH.split('\n');
    for (const linea of lineas) {
      const lineaClean = linea.replace(/^[•\-\s]+/, '').trim();
      const matchTipo = lineaClean.match(/^(T[1-4]):\s*(.+)/i);
      if (!matchTipo) continue;

      const tipo   = matchTipo[1].toUpperCase();
      const nombre = limpiarNombre(matchTipo[2].trim());
      if (!nombre) continue;

      // Generar ID único
      if (!contadorPorSemana[claveUS]) contadorPorSemana[claveUS] = 0;
      contadorPorSemana[claveUS]++;
      const idRecurso = `${claveUS}_R${contadorPorSemana[claveUS]}`;

      // Inferir tipo descriptivo
      const tipoDescriptivo = inferirTipo(tipo, nombre);

      recursos.push({
        unidad:   unidadNorm,
        semana:   semanaNorm,
        id:       idRecurso,
        nombre:   nombre,
        tipo:     tipoDescriptivo,
        estado:   'Sin comenzar',
        momento:  momento,
      });
    }
  }

  return recursos;
}

// ── Normalizar unidad ─────────────────────────────────────────────────────────

function normalizarUnidad(raw) {
  const r = raw.toLowerCase();
  if (r.includes('inicio') || r.includes('s0') || r === '0') return 'S0';
  // Buscar número romano o arábigo
  const romano = { 'i': 1, 'ii': 2, 'iii': 3, 'iv': 4, 'v': 5 };
  for (const [k, v] of Object.entries(romano)) {
    if (r.includes(' ' + k) || r.endsWith(' ' + k) || r.includes('unidad ' + k)) return `U${v}`;
  }
  const numMatch = r.match(/(\d+)/);
  if (numMatch) return `U${numMatch[1]}`;
  return 'U?';
}

// ── Normalizar semana ─────────────────────────────────────────────────────────

function normalizarSemana(raw) {
  const r = raw.toString().trim();
  if (/^s\d/i.test(r)) return r.toUpperCase();
  const numMatch = r.match(/(\d+)/);
  if (numMatch) return `S${numMatch[1]}`;
  return `S${r}`;
}

// ── Limpiar nombre del recurso ────────────────────────────────────────────────

function limpiarNombre(nombre) {
  // Quitar prefijos como "Presentación S1 —" si ya tienen el nombre
  // Mantener todo pero limpiar caracteres especiales al inicio
  return nombre
    .replace(/^[—–\-:]+\s*/, '')
    .replace(/Disponible en:.*$/i, '')
    .trim()
    .replace(/\s+/g, ' ')
    .substring(0, 120); // máx 120 caracteres
}

// ── Inferir tipo descriptivo ──────────────────────────────────────────────────

function inferirTipo(tipoT, nombre) {
  const nombreLower = nombre.toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, ''); // quitar acentos

  const mapa = TIPO_MAP[tipoT] || {};

  for (const [clave, valor] of Object.entries(mapa)) {
    if (clave === '_default') continue;
    if (nombreLower.includes(clave)) return valor;
  }

  return mapa['_default'] || tipoT;
}

// ── Insertar recursos en el seguimiento ───────────────────────────────────────

function insertarEnSeguimiento(recursos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const datos = hoja.getDataRange().getValues();

  // Recopilar IDs ya existentes en el seguimiento (col C, índice 2)
  const idsExistentes = new Set();
  for (const fila of datos) {
    const id = String(fila[CONFIG.COL_ID] || '').trim();
    if (id) idsExistentes.add(id);
  }

  // Encontrar la última fila con datos
  let ultimaFila = datos.length;
  while (ultimaFila > 0 && datos[ultimaFila - 1].every(c => !c)) {
    ultimaFila--;
  }

  // Agrupar recursos por unidad para insertar con encabezados de sección
  let filasAgregadas = 0;
  let filasDuplicadas = 0;
  const resumenTipos = {};

  // Determinar unidad actual al final del sheet para saber si hay que agregar encabezado
  let unidadActualEnSheet = '';
  for (let i = datos.length - 1; i >= 0; i--) {
    if (datos[i][CONFIG.COL_UNIDAD]) {
      unidadActualEnSheet = String(datos[i][CONFIG.COL_UNIDAD]);
      break;
    }
  }

  let filaInsercion = ultimaFila + 1;

  for (const rec of recursos) {
    // Verificar si ya existe
    if (idsExistentes.has(rec.id)) {
      filasDuplicadas++;
      continue;
    }

    // Insertar fila en el sheet
    const filaVacia = new Array(CONFIG.TOTAL_COLS).fill('');
    filaVacia[CONFIG.COL_UNIDAD] = rec.unidad;
    filaVacia[CONFIG.COL_SEMANA] = rec.semana;
    filaVacia[CONFIG.COL_ID]     = rec.id;
    filaVacia[CONFIG.COL_NOMBRE] = rec.nombre;
    filaVacia[CONFIG.COL_TIPO]   = rec.tipo;
    filaVacia[CONFIG.COL_ESTADO] = rec.estado;
    filaVacia[CONFIG.COL_RESP]   = CONFIG.DI_EMAIL;

    hoja.getRange(filaInsercion, 1, 1, CONFIG.TOTAL_COLS).setValues([filaVacia]);

    filasAgregadas++;
    filaInsercion++;
    idsExistentes.add(rec.id); // evitar duplicar en la misma corrida

    // Conteo por tipo
    resumenTipos[rec.tipo] = (resumenTipos[rec.tipo] || 0) + 1;
  }

  // Construir texto de resumen
  const resumen = Object.entries(resumenTipos)
    .map(([tipo, count]) => `  • ${tipo}: ${count}`)
    .join('\n');

  return { agregadas: filasAgregadas, duplicadas: filasDuplicadas, resumen };
}

// ── Ver resumen ───────────────────────────────────────────────────────────────

function verResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const datos = hoja.getDataRange().getValues();

  const conteo = {};
  let total = 0;

  for (let i = 4; i < datos.length; i++) {
    const tipo = String(datos[i][CONFIG.COL_TIPO] || '').trim();
    const estado = String(datos[i][CONFIG.COL_ESTADO] || '').trim();
    if (!tipo || tipo.startsWith('◀') || tipo === 'Tipo de recurso') continue;
    conteo[estado] = (conteo[estado] || 0) + 1;
    total++;
  }

  const resumen = Object.entries(conteo)
    .sort((a, b) => b[1] - a[1])
    .map(([estado, n]) => `  ${estado}: ${n}`)
    .join('\n');

  SpreadsheetApp.getUi().alert(
    '📊 Resumen del seguimiento',
    `Total de recursos: ${total}\n\nPor estado:\n${resumen}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
