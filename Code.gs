/**
 * ============================================================================
 * SISTEMA MES TEFLU — TRAZABILIDAD DE PRODUCCIÓN Y KANBAN FIP-55 REV. 3
 * Plataforma: Google Apps Script + Google Sheets
 * Cuenta Oficial: abteflu.mes@gmail.com
 * Spreadsheet ID: 1_QTYlxbj8rD9Jy7VwCiFarc5aEZbWiZg_2ujSNj38D4
 * ============================================================================
 */

const SPREADSHEET_ID = '1_QTYlxbj8rD9Jy7VwCiFarc5aEZbWiZg_2ujSNj38D4';

/**
 * Punto de entrada de la Web App
 */
function doGet(e) {
  try {
    ensureDatabaseInitialized_();
  } catch (err) {
    Logger.log('doGet ensureDatabaseInitialized_ warn: ' + err);
  }
  const template = HtmlService.createTemplateFromFile('Index');
  template.scanId = (e && e.parameter && e.parameter.kanban) ? e.parameter.kanban : '';
  return template.evaluate()
    .setTitle('MES TEFLU — Trazabilidad y Control de Producción')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/**
 * Obtiene la instancia del Spreadsheet Maestro
 */
function getDb_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Garantiza que las hojas operativas existan (los usuarios se leen 100% directo de CAT_USUARIOS_NOTIF)
 */
function ensureDatabaseInitialized_() {
  const ss = getDb_();

  // 0. MES_PROGRAMACION_SEMANAL (Programa Semanal Maestro)
  let wsProg = ss.getSheetByName('MES_PROGRAMACION_SEMANAL');
  if (!wsProg) {
    wsProg = ss.insertSheet('MES_PROGRAMACION_SEMANAL');
    wsProg.appendRow([
      'ID_PROGRAMA', 'SEMANA', 'FECHA_TURNO', 'ID_TURNO', 'HOJA_DE_RUTA', 'CLIENTE',
      'NUMERO_PARTE', 'OPERACION', 'MAQUINA_ASIGNADA', 'TIPO_MAQUINA', 'RATE_SEG',
      'CODIGO_MP_USADO', 'ES_MP_ALTERNA', 'FOLIO_DESVIACION', 'LONGITUD_BARRA_MM',
      'CORTES_POR_BARRA', 'BARRAS_A_OCUPAR', 'PIEZAS_PROGRAMADAS', 'PIEZAS_TOMADAS_WIP',
      'PIEZAS_ENTREGADAS_OK', 'PIEZAS_A_WIP', 'PIEZAS_SCRAP', 'PORCENTAJE_CUMPLIMIENTO',
      'ESTADO_ORDEN', 'CREADO_POR', 'FECHA_CREACION', 'ULTIMA_MODIFICACION_POR',
      'FECHA_ULTIMA_MODIFICACION', 'COMENTARIOS'
    ]);
  }

  // 1. MES_LOTES_KANBAN (Tarjetas de Contenedores FIP-55 Rev. 3 de Piso)
  let wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
  if (!wsKanban) {
    wsKanban = ss.insertSheet('MES_LOTES_KANBAN');
    wsKanban.appendRow([
      'ID_KANBAN', 'ID_PROGRAMA', 'HOJA_DE_RUTA', 'NUMERO_PARTE', 'CLIENTE',
      'CODIGO_MP', 'LOTE_FISICO_MP', 'ES_MP_ALTERNA', 'FOLIO_DESVIACION',
      'CANTIDAD_INICIAL', 'CANTIDAD_ACTUAL', 'OPERACION_ACTUAL', 'MAQUINA_ACTUAL',
      'ESTADO_KANBAN', 'SIGUIENTE_OPERACION', 'AUDITADO_FISICO',
      'FECHA_ULTIMA_AUDITORIA', 'AUDITADO_POR', 'FECHA_CREACION', 'ES_SEGURIDAD'
    ]);
  }

  // 2. MES_INVENTARIO_WIP (Stock en Proceso / Pulmón)
  let wsWip = ss.getSheetByName('MES_INVENTARIO_WIP');
  if (!wsWip) {
    wsWip = ss.insertSheet('MES_INVENTARIO_WIP');
    wsWip.appendRow([
      'ID_WIP', 'ID_KANBAN_ORIGEN', 'HOJA_DE_RUTA', 'NUMERO_PARTE', 'CLIENTE',
      'OPERACION_ORIGEN', 'SIGUIENTE_PROCESO', 'MAQUINA_ORIGEN', 'CODIGO_MP',
      'PIEZAS_EN_WIP', 'UBICACION_PLANTA', 'ESTATUS_AUDITORIA', 'FECHA_AUDITORIA',
      'AUDITADO_POR', 'FECHA_CREACION', 'ESTADO_WIP'
    ]);
  }

  // 3. MES_AUDITORIA_FISICA (Conciliaciones de Piso)
  let wsAudit = ss.getSheetByName('MES_AUDITORIA_FISICA');
  if (!wsAudit) {
    wsAudit = ss.insertSheet('MES_AUDITORIA_FISICA');
    wsAudit.appendRow([
      'ID_AUDITORIA', 'FECHA', 'USUARIO', 'TIPO_INVENTARIO', 'ID_REGISTRO',
      'NUMERO_PARTE', 'CANTIDAD_SISTEMA', 'CANTIDAD_FISICA', 'DIFERENCIA',
      'RESULTADO', 'MOTIVO'
    ]);
  }

  // 4. MES_SCRAP_REGISTROS (Registro de Scrap Auditado)
  let wsScrap = ss.getSheetByName('MES_SCRAP_REGISTROS');
  if (!wsScrap) {
    wsScrap = ss.insertSheet('MES_SCRAP_REGISTROS');
    wsScrap.appendRow([
      'ID_SCRAP', 'FECHA', 'SEMANA', 'HOJA_DE_RUTA', 'ID_KANBAN', 'CLIENTE',
      'NUMERO_PARTE', 'OPERACION', 'MAQUINA', 'TURNO', 'SUPERVISOR', 'OPERADOR',
      'CODIGO_MP', 'COLADA_MP', 'PIEZAS_SCRAP', 'MODO_DE_FALLA', 'DEFECTO',
      'CAUSA', 'DEPARTAMENTO', 'AUDITADO_POR', 'COMENTARIOS'
    ]);
  }

  // 5. MES_BITACORA_CAMBIOS (Trazabilidad y Auditoría)
  let wsBit = ss.getSheetByName('MES_BITACORA_CAMBIOS');
  if (!wsBit) {
    wsBit = ss.insertSheet('MES_BITACORA_CAMBIOS');
    wsBit.appendRow([
      'ID_CAMBIO', 'FECHA', 'USUARIO', 'ID_PROGRAMA', 'HOJA_DE_RUTA',
      'NUMERO_PARTE', 'CAMPO', 'VALOR_ANTERIOR', 'VALOR_NUEVO', 'MOTIVO'
    ]);
  }

  // 6. MES_NOTIFICACIONES_LOG (Registro de Alertas y WhatsApp)
  let wsNotif = ss.getSheetByName('MES_NOTIFICACIONES_LOG');
  if (!wsNotif) {
    wsNotif = ss.insertSheet('MES_NOTIFICACIONES_LOG');
    wsNotif.appendRow([
      'ID_NOTIFICACION', 'FECHA', 'TIPO_EVENTO', 'USUARIO_ORIGEN',
      'DESTINATARIOS', 'CANAL', 'MENSAJE', 'ESTATUS'
    ]);
  }

  // 7. MES_APORTACIONES_PISO (Genealogía: quién cortó/procesó, en qué máquina, con qué MP y cuándo)
  let wsApt = ss.getSheetByName('MES_APORTACIONES_PISO');
  if (!wsApt) {
    wsApt = ss.insertSheet('MES_APORTACIONES_PISO');
    wsApt.appendRow([
      'ID_APORTACION', 'FECHA_HORA', 'ID_KANBAN', 'ID_PROGRAMA', 'HOJA_DE_RUTA',
      'NUMERO_PARTE', 'CLIENTE', 'OPERACION', 'MAQUINA', 'SLOT',
      'NUM_EMPLEADO', 'NOMBRE_OPERADOR', 'TURNO', 'PIEZAS_OK', 'PIEZAS_WIP',
      'CONTENEDOR_CERRADO', 'CODIGO_MP', 'COLADA_MP', 'USUARIO_TERMINAL'
    ]);
  }

  // 8. Migración no destructiva: columnas adicionales requeridas por la captura de piso
  ensureColumns_(ss.getSheetByName('MES_LOTES_KANBAN'), [
    'CANTIDAD_ESTANDAR_PACK', 'LOTE_CONTENEDOR', 'TIPO_EMPAQUE', 'NOMBRE_EMPAQUE',
    'ES_PACK_MANUAL', 'CONSECUTIVO', 'O1_JSON', 'O2_JSON', 'O3_JSON', 'TRAZABILIDAD_ORIGEN'
  ]);
  ensureColumns_(ss.getSheetByName('MES_SCRAP_REGISTROS'), [
    'TIPO_RACK', 'ESTATUS', 'NUM_EMPLEADO', 'ID_PROGRAMA', 'CANT_RECUPERADA', 'CANT_DESCARTADA'
  ]);
}

/**
 * Agrega al final de la fila 1 las columnas que falten (sin mover ni borrar datos existentes)
 */
function ensureColumns_(ws, cols) {
  if (!ws) return;
  try {
    const lastCol = Math.max(1, ws.getLastColumn());
    const headers = ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
    const missing = cols.filter(function(c) { return headers.indexOf(c) === -1; });
    if (!missing.length) return;
    let used = headers.length;
    while (used > 0 && headers[used - 1] === '') used--;
    if (ws.getMaxColumns() < used + missing.length) {
      ws.insertColumnsAfter(ws.getMaxColumns(), used + missing.length - ws.getMaxColumns());
    }
    ws.getRange(1, used + 1, 1, missing.length).setValues([missing]);
  } catch (e) {
    Logger.log('ensureColumns_ warn (' + ws.getName() + '): ' + e);
  }
}

/**
 * Convierte un objeto {ENCABEZADO: valor} en una fila alineada a los encabezados reales de la hoja
 */
function rowByHeaders_(headers, obj) {
  return headers.map(function(h) {
    return (h && obj[h] !== undefined && obj[h] !== null) ? obj[h] : '';
  });
}

function appendByHeaders_(ws, obj) {
  const lastCol = ws.getLastColumn();
  const headers = ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  ws.appendRow(rowByHeaders_(headers, obj));
}

/**
 * Ejecuta una operación de escritura con candado para evitar que dos terminales
 * (operadores simultáneos) se pisen al actualizar la misma orden o tarjeta.
 */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function uniqueId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss') + '-' + Utilities.getUuid().substring(0, 4).toUpperCase();
}

/**
 * Convierte una hoja en arreglo de objetos JSON usando la fila 1 como claves
 */
function sheetToObjects_(sheetName) {
  const ss = getDb_();
  const ws = ss.getSheetByName(sheetName);
  if (!ws) return [];
  const maxCols = ws.getMaxColumns();
  const lastCol = Math.min(maxCols, Math.max(ws.getLastColumn(), sheetName === 'CAT_USUARIOS_NOTIF' ? 14 : (sheetName === 'CAT_NUMEROS_PARTE' ? 22 : 1)));
  const lastRow = ws.getLastRow();
  if (lastRow < 2 || lastCol < 1) return [];
  const values = ws.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(h => String(h).trim());
  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(cell => String(cell).trim() === '')) continue;
    const obj = { _rowIndex: i + 1 };
    headers.forEach((h, idx) => {
      if (h) obj[h] = row[idx];
    });
    // Soporte para Columna N (índice 13 en base 0: Col A=0, ..., Col N=13) en CAT_USUARIOS_NOTIF
    if (sheetName === 'CAT_USUARIOS_NOTIF' || row.length >= 14) {
      const valColN = String(row[13] || '').trim();
      const hdrColN = String(headers[13] || '').trim().toUpperCase();
      const isBooleanFlag = (valColN.toUpperCase() === 'SI' || valColN.toUpperCase() === 'NO') && hdrColN.indexOf('RECIBE_NOTIF') !== -1;
      if (valColN && !isBooleanFlag) {
        obj.COLUMNA_N = valColN;
        obj.PERSONA = valColN;
      }
    }
    // Soporte para Columnas O (14), P (15), Q (16), R (17) en CAT_NUMEROS_PARTE (Standard Packs por proceso)
    if (sheetName === 'CAT_NUMEROS_PARTE') {
      const valColO = String(row[14] || '').trim();
      const valColP = String(row[15] || '').trim();
      const valColQ = String(row[16] || '').trim();
      const valColR = String(row[17] || '').trim();
      if (valColO) {
        obj.COLUMNA_O = valColO;
        obj.PACK_CORTE = valColO;
        obj.STANDAR_PACK_CORTE = valColO;
        obj.STANDAR_PACK_PRINCIPAL = valColO;
        obj.STANDAR_PACK = valColO;
      }
      if (valColP) {
        obj.COLUMNA_P = valColP;
        obj.PACKALT_CORTE = valColP;
        obj.STANDAR_PACKALT_CORTE = valColP;
        obj.STANDAR_PACK_ALTERNO = valColP;
      }
      if (valColQ) {
        obj.COLUMNA_Q = valColQ;
        obj.PACK_DOBLADO = valColQ;
        obj.STANDAR_PACK_DOBLADO = valColQ;
      }
      if (valColR) {
        obj.COLUMNA_R = valColR;
        obj.PACKALT_DOBLADO = valColR;
        obj.STANDAR_PACKALT_DOBLADO = valColR;
      }
    }
    result.push(obj);
  }
  return result;
}

/**
 * Autenticación de usuario consultando en tiempo real la hoja CAT_USUARIOS_NOTIF
 */
function loginUser(username, password) {
  const users = sheetToObjects_('CAT_USUARIOS_NOTIF');
  const uClean = String(username || '').trim().toLowerCase();
  const pClean = String(password || '').trim();

  const found = users.find(u =>
    String(u.USUARIO || '').trim().toLowerCase() === uClean &&
    String(u.PASSWORD || '').trim() === pClean &&
    String(u.ACTIVO || 'SI').toUpperCase() === 'SI'
  );

  if (!found) {
    return { ok: false, message: 'Usuario o contraseña incorrectos, o cuenta inactiva en CAT_USUARIOS_NOTIF.' };
  }

  // Nombre de la persona asignada al usuario (Columna N de CAT_USUARIOS_NOTIF)
  const personaColN = String(found.COLUMNA_N || found.PERSONA || found.NOMBRE_PERSONA || found.COLABORADOR || found.RESPONSABLE || '').trim();

  return {
    ok: true,
    user: {
      id: found.ID_USUARIO,
      username: found.USUARIO,
      nombre: found.NOMBRE_COMPLETO,
      rol: found.ROL,
      persona: personaColN,
      colN: personaColN,
      whatsapp: found.TELEFONO_WHATSAPP || ''
    }
  };
}

/**
 * Permite al usuario Master (ADMIN / MASTER) crear o modificar usuarios directamente en CAT_USUARIOS_NOTIF
 */
function saveUserInCatalog(payload) {
  // payload: { adminRol, adminUser, idUsuario, usuario, password, nombreCompleto, email, rol, whatsapp, activo }
  const r = String(payload.adminRol || '').trim().toUpperCase();
  if (r !== 'ADMIN' && r !== 'MASTER') {
    throw new Error('Acceso denegado: Solo el usuario Master puede administrar usuarios.');
  }
  const ss = getDb_();
  const ws = ss.getSheetByName('CAT_USUARIOS_NOTIF');
  if (!ws) throw new Error('No se encontró la hoja CAT_USUARIOS_NOTIF.');

  const data = ws.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxId = headers.indexOf('ID_USUARIO');
  const idxUser = headers.indexOf('USUARIO');
  const idxPass = headers.indexOf('PASSWORD');
  const idxNombre = headers.indexOf('NOMBRE_COMPLETO');
  const idxEmail = headers.indexOf('EMAIL');
  const idxRol = headers.indexOf('ROL');
  const idxWa = headers.indexOf('TELEFONO_WHATSAPP');
  const idxActivo = headers.indexOf('ACTIVO');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (
      (payload.idUsuario && String(data[i][idxId]).trim() === String(payload.idUsuario).trim()) ||
      String(data[i][idxUser]).trim().toLowerCase() === String(payload.usuario).trim().toLowerCase()
    ) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow !== -1) {
    if (idxUser >= 0) ws.getRange(targetRow, idxUser + 1).setValue(String(payload.usuario).trim());
    if (idxPass >= 0) ws.getRange(targetRow, idxPass + 1).setValue(String(payload.password).trim());
    if (idxNombre >= 0) ws.getRange(targetRow, idxNombre + 1).setValue(String(payload.nombreCompleto || '').trim());
    if (idxEmail >= 0) ws.getRange(targetRow, idxEmail + 1).setValue(String(payload.email || 'abteflu.mes@gmail.com').trim());
    if (idxRol >= 0) ws.getRange(targetRow, idxRol + 1).setValue(String(payload.rol || 'OPERADOR').trim().toUpperCase());
    if (idxWa >= 0) ws.getRange(targetRow, idxWa + 1).setValue(String(payload.whatsapp || '').trim());
    if (idxActivo >= 0) ws.getRange(targetRow, idxActivo + 1).setValue(String(payload.activo || 'SI').trim().toUpperCase());
    if (payload.persona !== undefined && String(payload.persona).trim() !== '') {
      const idxPersona = headers.indexOf('PERSONA');
      if (idxPersona >= 0) ws.getRange(targetRow, idxPersona + 1).setValue(String(payload.persona).trim());
      else ws.getRange(targetRow, 14).setValue(String(payload.persona).trim());
    }
  } else {
    const newId = payload.idUsuario || ('USR-' + String(data.length).padStart(3, '0'));
    const rolUpper = String(payload.rol || 'OPERADOR').trim().toUpperCase();
    appendByHeaders_(ws, {
      ID_USUARIO: newId,
      USUARIO: String(payload.usuario).trim(),
      PASSWORD: String(payload.password).trim(),
      NOMBRE_COMPLETO: String(payload.nombreCompleto || '').trim(),
      DEPARTAMENTO: rolUpper,
      EMAIL: String(payload.email || 'abteflu.mes@gmail.com').trim(),
      ROL: rolUpper,
      NIVEL_PRIORIDAD: 5,
      TELEFONO_WHATSAPP: String(payload.whatsapp || '').trim(),
      RECIBE_NOTIF_INGENIERIA: 'SI',
      RECIBE_NOTIF_PROGRAMACION: 'SI',
      RECIBE_NOTIF_MODIFICACIONES: 'SI',
      RECIBE_NOTIF_DESVIACION_MP: 'NO',
      RECIBE_NOTIF_AUDITORIA_WIP: String(payload.persona || 'SI').trim(),
      RECIBE_NOTIF_FLUJO_KANBAN: 'SI',
      PERSONA: String(payload.persona || '').trim(),
      ACTIVO: String(payload.activo || 'SI').trim().toUpperCase()
    });
    if (payload.persona && headers.indexOf('PERSONA') === -1) {
      ws.getRange(ws.getLastRow(), 14).setValue(String(payload.persona).trim());
    }
  }

  logBitacora_({
    usuario: payload.adminUser || 'master',
    idPrograma: 'CAT_USUARIOS_NOTIF',
    hojaDeRuta: 'N/A',
    numeroParte: 'N/A',
    campo: 'GESTION_USUARIO (' + payload.usuario + ')',
    valorAnterior: targetRow !== -1 ? 'ACTUALIZACION' : 'NUEVO',
    valorNuevo: 'ROL: ' + payload.rol + ' | ACTIVO: ' + (payload.activo || 'SI'),
    motivo: 'Modificación desde módulo Master'
  });

  return {
    ok: true,
    usuarios: sheetToObjects_('CAT_USUARIOS_NOTIF')
  };
}

/**
 * Obtiene los segundos efectivos de cada turno directamente desde la hoja CAT_TURNOS
 */
function getShiftSecondsMap_() {
  const turnos = sheetToObjects_('CAT_TURNOS');
  const map = {};
  turnos.forEach(t => {
    const id = String(t.ID_TURNO || '').trim().toUpperCase();
    if (!id) return;
    const minAct = Number(t.MINUTOS_ACTIVOS || 0);
    const segEf = Number(t.SEGUNDOS_EFECTIVOS || (minAct > 0 ? minAct * 60 : 0));
    if (segEf > 0) map[id] = segEf;
  });
  return map;
}

function getShiftWeight_(turno) {
  const t = String(turno || '').trim().toUpperCase();
  if (t === 'T1' || t.indexOf('1') === 0) return 1;
  if (t === 'T2' || t.indexOf('2') === 0) return 2;
  if (t === 'T3' || t.indexOf('3') === 0) return 3;
  return 9;
}

function sortProgramacionChronology_(list) {
  return (list || []).sort(function(a, b) {
    const fa = normalizeIsoDate_(a.FECHA_TURNO || a.FECHA || '');
    const fb = normalizeIsoDate_(b.FECHA_TURNO || b.FECHA || '');
    if (fa !== fb) return fa.localeCompare(fb);
    const wa = getShiftWeight_(a.ID_TURNO || a.TURNO);
    const wb = getShiftWeight_(b.ID_TURNO || b.TURNO);
    if (wa !== wb) return wa - wb;
    const ma = String(a.MAQUINA_ASIGNADA || a.MAQUINA || '');
    const mb = String(b.MAQUINA_ASIGNADA || b.MAQUINA || '');
    if (ma !== mb) return ma.localeCompare(mb);
    return String(a.HOJA_DE_RUTA || '').localeCompare(String(b.HOJA_DE_RUTA || ''));
  });
}

function toNumSafe_(val) {
  if (typeof val === 'number') return isFinite(val) ? val : 0;
  const cleaned = String(val === null || val === undefined ? '' : val).replace(/,/g, '').trim();
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return isFinite(n) ? n : 0;
}

function sheetToObjectsSafe_(sheetName) {
  try {
    return sheetToObjects_(sheetName);
  } catch (e) {
    Logger.log('sheetToObjectsSafe_ error en "' + sheetName + '": ' + e);
    return [];
  }
}

function parseSlotJson_(val) {
  const empty = { nombre: '', emp: '', cant: '', fecha: '' };
  if (!val) return empty;
  try {
    const parsed = JSON.parse(String(val));
    const cantRaw = parsed.cant;
    const cantClean = (cantRaw !== undefined && cantRaw !== null && cantRaw !== '')
      ? (isFinite(Number(String(cantRaw).replace(/,/g, ''))) ? Number(String(cantRaw).replace(/,/g, '')) : String(cantRaw))
      : '';
    return {
      nombre: String(parsed.nombre || ''),
      emp: String(parsed.emp || ''),
      cant: cantClean,
      fecha: String(parsed.fecha || '')
    };
  } catch (e) {
    return empty;
  }
}

/**
 * Carga inicial de todos los catálogos y estado operativo de la planta
 */
function getInitialData() {
  let initWarning = '';
  try {
    ensureDatabaseInitialized_();
  } catch (initErr) {
    initWarning = String(initErr && initErr.message ? initErr.message : initErr);
    Logger.log('getInitialData ensureDatabaseInitialized_ warn: ' + initWarning);
  }

  const props = PropertiesService.getScriptProperties();
  const modoAuditoriaExterna = props.getProperty('MODO_AUDITORIA_EXTERNA') === 'TRUE';
  const whatsappWebhookUrl = props.getProperty('WHATSAPP_WEBHOOK_URL') || '';

  const kanbans = sheetToObjectsSafe_('MES_LOTES_KANBAN').map(function(k) {
    k.CANTIDAD_ACTUAL = toNumSafe_(k.CANTIDAD_ACTUAL);
    k.CANTIDAD_INICIAL = toNumSafe_(k.CANTIDAD_INICIAL);
    if (k.CANTIDAD_ESTANDAR_PACK !== undefined && k.CANTIDAD_ESTANDAR_PACK !== '') {
      k.CANTIDAD_ESTANDAR_PACK = toNumSafe_(k.CANTIDAD_ESTANDAR_PACK);
    }
    if (k.ES_PACK_MANUAL !== undefined) {
      k.ES_PACK_MANUAL = (k.ES_PACK_MANUAL === true || String(k.ES_PACK_MANUAL).toUpperCase() === 'TRUE' || String(k.ES_PACK_MANUAL).toUpperCase() === 'SI');
    }
    k.O1 = parseSlotJson_(k.O1_JSON);
    k.O2 = parseSlotJson_(k.O2_JSON);
    k.O3 = parseSlotJson_(k.O3_JSON);
    return k;
  });

  const payload = {
    config: {
      modoAuditoriaExterna: modoAuditoriaExterna,
      whatsappConfigured: Boolean(whatsappWebhookUrl),
      spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit',
      initWarning: initWarning
    },
    catalogs: {
      partes: sheetToObjectsSafe_('CAT_NUMEROS_PARTE'),
      rutas: sheetToObjectsSafe_('CAT_RUTAS_TIEMPOS'),
      materiaPrima: sheetToObjectsSafe_('CAT_MATERIA_PRIMA'),
      maquinas: sheetToObjectsSafe_('CAT_MAQUINAS'),
      herramentales: sheetToObjectsSafe_('CAT_HERRAMENTALES'),
      turnos: sheetToObjectsSafe_('CAT_TURNOS'),
      defectos: sheetToObjectsSafe_('CAT_DEFECTOS_SCRAP'),
      paros: sheetToObjectsSafe_('CAT_CAUSAS_PARO'),
      usuarios: sheetToObjectsSafe_('CAT_USUARIOS_NOTIF')
    },
    operational: {
      programacion: sortProgramacionChronology_(sheetToObjectsSafe_('MES_PROGRAMACION_SEMANAL')),
      lotesKanban: kanbans,
      inventarioWip: sheetToObjectsSafe_('MES_INVENTARIO_WIP'),
      auditoriaFisica: sheetToObjectsSafe_('MES_AUDITORIA_FISICA'),
      scrapRegistros: sheetToObjectsSafe_('MES_SCRAP_REGISTROS'),
      aportacionesPiso: sheetToObjectsSafe_('MES_APORTACIONES_PISO').slice(-300).reverse(),
      bitacoraCambios: sheetToObjectsSafe_('MES_BITACORA_CAMBIOS').slice(-100).reverse(),
      notificacionesLog: sheetToObjectsSafe_('MES_NOTIFICACIONES_LOG').slice(-50).reverse()
    }
  };

  // Garantiza serialización 100% compatible con google.script.run (elimina cualquier NaN/Date/undefined que causaría retorno null)
  return JSON.parse(JSON.stringify(payload));
}

/**
 * ============================================================================
 * MÓDULO 0: INGENIERÍA — CAMBIO DE MÁQUINA DE PRIORIDAD (P <-> O) Y RATES
 * ============================================================================
 */
function updateEngineeringRoute(payload) {
  // payload: { usuario, rol, numeroParte, operacion, nuevaMaquinaPrincipal, nuevoRateSeg, esPiezaSeguridad, motivo }
  const r = String(payload.rol || '').trim().toUpperCase();
  if (r !== 'INGENIERIA' && r !== 'DOC_TECNICA' && r !== 'ADMIN' && r !== 'MASTER') {
    throw new Error('Acceso denegado: Solo INGENIERÍA o MASTER puede modificar máquinas principales (P) o Rates.');
  }

  const ss = getDb_();
  const wsRutas = ss.getSheetByName('CAT_RUTAS_TIEMPOS');
  const data = wsRutas.getDataRange().getValues();
  const headers = data[0];
  const shiftSec = getShiftSecondsMap_();

  const colNP = headers.indexOf('NUMERO_PARTE');
  const colOp = headers.indexOf('OPERACION');
  const colMaq = headers.indexOf('MAQUINA');
  const colTipo = headers.indexOf('TIPO_MAQUINA');
  const colRate = headers.indexOf('RATE_SEGUNDOS');
  const colPzasHr = headers.indexOf('PIEZAS_POR_HORA');
  const colCapT1 = headers.findIndex(h => String(h).indexOf('CAPACIDAD_T1') === 0);
  const colCapT2 = headers.findIndex(h => String(h).indexOf('CAPACIDAD_T2') === 0);
  const colCapT3 = headers.findIndex(h => String(h).indexOf('CAPACIDAD_T3') === 0);
  const colCapSab = headers.findIndex(h => String(h).indexOf('CAPACIDAD_SAB') === 0);
  const colObs = headers.indexOf('OBSERVACIONES');

  let anteriorPrincipal = 'Ninguna';
  let encontrado = false;

  for (let i = 1; i < data.length; i++) {
    const np = String(data[i][colNP]).trim();
    const op = String(data[i][colOp]).trim();
    const maq = String(data[i][colMaq]).trim();
    const tipoActual = String(data[i][colTipo]).trim().toUpperCase();

    if (np === payload.numeroParte && op === payload.operacion) {
      if (tipoActual === 'P' && maq !== payload.nuevaMaquinaPrincipal) {
        anteriorPrincipal = maq;
        wsRutas.getRange(i + 1, colTipo + 1).setValue('O');
      }
      if (maq === payload.nuevaMaquinaPrincipal) {
        encontrado = true;
        wsRutas.getRange(i + 1, colTipo + 1).setValue('P');
        if (payload.nuevoRateSeg && Number(payload.nuevoRateSeg) > 0) {
          const r = Number(payload.nuevoRateSeg);
          wsRutas.getRange(i + 1, colRate + 1).setValue(r);
          wsRutas.getRange(i + 1, colPzasHr + 1).setValue(Math.round((3600 / r) * 10) / 10);
          if (colCapT1 >= 0 && shiftSec['T1']) wsRutas.getRange(i + 1, colCapT1 + 1).setValue(Math.floor(shiftSec['T1'] / r));
          if (colCapT2 >= 0 && shiftSec['T2']) wsRutas.getRange(i + 1, colCapT2 + 1).setValue(Math.floor(shiftSec['T2'] / r));
          if (colCapT3 >= 0 && shiftSec['T3']) wsRutas.getRange(i + 1, colCapT3 + 1).setValue(Math.floor(shiftSec['T3'] / r));
          if (colCapSab >= 0 && (shiftSec['T1_SAB'] || shiftSec['T2_SAB'])) {
            wsRutas.getRange(i + 1, colCapSab + 1).setValue(Math.floor((shiftSec['T1_SAB'] || shiftSec['T2_SAB']) / r));
          }
        }
      }
      if (colObs >= 0) {
        wsRutas.getRange(i + 1, colObs + 1).setValue('');
      }
    }
  }

  if (!encontrado) {
    throw new Error('No se encontró la combinación de Número de Parte, Operación y Máquina en el catálogo.');
  }

  logBitacora_({
    usuario: payload.usuario,
    idPrograma: 'CATALOGO-ING',
    hojaDeRuta: 'N/A',
    numeroParte: payload.numeroParte,
    campo: 'MAQUINA_PRINCIPAL (' + payload.operacion + ')',
    valorAnterior: anteriorPrincipal,
    valorNuevo: payload.nuevaMaquinaPrincipal + (payload.nuevoRateSeg ? ' (' + payload.nuevoRateSeg + 's)' : ''),
    motivo: payload.motivo || 'Actualización de prioridad por Ingeniería'
  });

  sendWhatsAppNotification_({
    tipoEvento: 'INGENIERIA_CAMBIO_PRIORIDAD',
    usuarioOrigen: payload.usuario,
    destinatarios: 'PROGRAMACION, SUPERVISOR, INGENIERIA',
    mensaje: '🔧 [INGENIERÍA MES] Cambio de máquina Principal (P) en N/P ' + payload.numeroParte +
             ' (' + payload.operacion + '): Anterior=' + anteriorPrincipal +
             ' -> Nueva Principal=' + payload.nuevaMaquinaPrincipal +
             '. Motivo: ' + (payload.motivo || 'Ajuste de ingeniería')
  });

  return { ok: true, message: 'Prioridad de máquina actualizada con éxito. Relevo P <-> O aplicado.' };
}

/**
 * ============================================================================
 * MÓDULO 1: PROGRAMACIÓN SEMANAL (DOM 22:00 A SÁB 16:30, DESVIACIÓN MP, WIP)
 * ============================================================================
 */
function normalizeIsoDate_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(val).trim().split(' ')[0];
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    return m[3] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');
  }
  return s;
}

function createOrUpdateProgramacion(payload) {
  return withLock_(function() {
  const rolUpper = String(payload.rol || '').trim().toUpperCase();
  if (rolUpper !== 'PROGRAMACION' && rolUpper !== 'ADMIN' && rolUpper !== 'MASTER') {
    throw new Error('Acceso denegado: Solo el usuario de PROGRAMACIÓN o Master puede detonar o modificar el programa.');
  }

  const ss = getDb_();
  const wsProg = ss.getSheetByName('MES_PROGRAMACION_SEMANAL');
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // Si es modificación de un programa existente
  if (payload.idProgramaExistente) {
    const data = wsProg.getDataRange().getValues();
    const headers = data[0];
    const idxId = headers.indexOf('ID_PROGRAMA');
    const idxPzasProg = headers.indexOf('PIEZAS_PROGRAMADAS');
    const idxPzasOk = headers.indexOf('PIEZAS_ENTREGADAS_OK');
    const idxMaq = headers.indexOf('MAQUINA_ASIGNADA');
    const idxColada = headers.indexOf('COMENTARIOS');
    const idxModPor = headers.indexOf('ULTIMA_MODIFICACION_POR');
    const idxFechaMod = headers.indexOf('FECHA_ULTIMA_MODIFICACION');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxId]) === String(payload.idProgramaExistente)) {
        const entregadasActuales = Number(data[i][idxPzasOk] || 0);
        const nuevasProgramadas = Number(payload.piezasProgramadas);

        // Candado: No bajar cantidad por debajo de lo ya producido en piso
        if (nuevasProgramadas < entregadasActuales) {
          throw new Error(
            'CANDADO DE PISO: No puedes bajar la cantidad programada a ' + nuevasProgramadas +
            ' pzas porque la máquina ya lleva ' + entregadasActuales + ' piezas producidas.'
          );
        }

        const valorAnt = 'Pzas: ' + data[i][idxPzasProg] + ' | Máq: ' + data[i][idxMaq];
        wsProg.getRange(i + 1, idxPzasProg + 1).setValue(nuevasProgramadas);
        const idxBarrasMod = headers.indexOf('BARRAS_A_OCUPAR');
        const idxCortesMod = headers.indexOf('CORTES_POR_BARRA');
        if (idxBarrasMod >= 0) {
          const cVal = Number(data[i][idxCortesMod] || payload.cortesPorBarra || 1) || 1;
          const bVal = Math.ceil(nuevasProgramadas / Math.max(1, Math.floor(cVal)));
          wsProg.getRange(i + 1, idxBarrasMod + 1).setValue(bVal);
        }
        if (payload.maquinaAsignada) {
          wsProg.getRange(i + 1, idxMaq + 1).setValue(payload.maquinaAsignada);
        }
        if (payload.coladaMp) {
          wsProg.getRange(i + 1, idxColada + 1).setValue('COLADA:' + payload.coladaMp + ' | ' + (payload.comentarios || ''));
        }
        wsProg.getRange(i + 1, idxModPor + 1).setValue(payload.usuario);
        wsProg.getRange(i + 1, idxFechaMod + 1).setValue(nowStr);

        logBitacora_({
          usuario: payload.usuario,
          idPrograma: payload.idProgramaExistente,
          hojaDeRuta: data[i][headers.indexOf('HOJA_DE_RUTA')],
          numeroParte: data[i][headers.indexOf('NUMERO_PARTE')],
          campo: 'MODIFICACION_PROGRAMA',
          valorAnterior: valorAnt,
          valorNuevo: 'Pzas: ' + nuevasProgramadas + ' | Máq: ' + (payload.maquinaAsignada || data[i][idxMaq]),
          motivo: payload.motivoCambio || 'Ajuste de programación semanal'
        });

        sendWhatsAppNotification_({
          tipoEvento: 'MODIFICACION_PROGRAMACION',
          usuarioOrigen: payload.usuario,
          destinatarios: 'TODOS',
          mensaje: '⚠️ [MODIFICACIÓN PROGRAMA] Hoja de Ruta #' + data[i][headers.indexOf('HOJA_DE_RUTA')] +
                   ' (N/P ' + data[i][headers.indexOf('NUMERO_PARTE')] + ') modificada por ' + payload.usuario +
                   '. Antes: [' + valorAnt + '] -> Ahora: [Pzas: ' + nuevasProgramadas + ']. Motivo: ' + payload.motivoCambio
        });

        return { ok: true, message: 'Programación modificada y registrada en bitácora.' };
      }
    }
    throw new Error('No se encontró el ID de Programa a modificar.');
  }

  // Creación de nueva orden de programación semanal
  if (payload.esMpAlterna === 'SI' && !payload.folioDesviacion) {
    throw new Error('Para utilizar Materia Prima Alterna (diferente longitud de barra) es obligatorio capturar el Folio de Desviación.');
  }

  // Candado de Capacidad Real por Turno y Máquina (Tiempo Restante en CAT_TURNOS)
  const shiftSecMap = getShiftSecondsMap_();
  const fIso = normalizeIsoDate_(payload.fechaTurno);
  const tClean = String(payload.idTurno || '').trim().toUpperCase();
  const mClean = String(payload.maquinaAsignada || '').trim();
  const opUpper = String(payload.operacion || '').trim().toUpperCase();
  const isCorteOp = opUpper.indexOf('CORTE') !== -1 || mClean.startsWith('CL-') || mClean === 'C-123' || mClean === 'C-114' || mClean === 'C-116';
  const isSabado = Boolean(fIso && new Date(fIso + 'T12:00:00').getDay() === 6);
  const totalSecTurno = Number(
    (isSabado && shiftSecMap[tClean + '_SAB']) ? shiftSecMap[tClean + '_SAB'] : (shiftSecMap[tClean] || 0)
  );
  const rateNueva = Number(payload.rateSeg || 0);
  const pzasEnMaquinaNueva = Math.max(0, Number(payload.piezasProgramadas || 0) - Number(payload.piezasTomadasDeWip || 0));

  // CANDADO DE CORTE PREVIO: Si es un proceso auxiliar (Doblado, Machuelado, Soldadura, etc.),
  // solo se puede programar si ya existe material cortado entregado en planta.
  if (!isCorteOp) {
    const progRows = sheetToObjects_('MES_PROGRAMACION_SEMANAL');
    const wipRows = sheetToObjects_('MES_INVENTARIO_WIP');
    const kanbanRows = sheetToObjects_('MES_LOTES_KANBAN');
    const npClean = String(payload.numeroParte || '').trim();

    let totalCortadoOk = 0;
    progRows.forEach(r => {
      const rNP = String(r.NUMERO_PARTE || '').trim();
      const rOp = String(r.OPERACION || '').trim().toUpperCase();
      const rMaq = String(r.MAQUINA_ASIGNADA || '').trim().toUpperCase();
      const rEst = String(r.ESTADO_ORDEN || '').trim().toUpperCase();
      const isRCorte = rOp.indexOf('CORTE') !== -1 || rMaq.startsWith('CL-') || rMaq === 'C-123' || rMaq === 'C-114' || rMaq === 'C-116';
      if (rNP === npClean && isRCorte && rEst !== 'CANCELADA') {
        totalCortadoOk += Number(r.PIEZAS_ENTREGADAS_OK || 0);
      }
    });

    let totalWipCortado = 0;
    wipRows.forEach(w => {
      const wNP = String(w.NUMERO_PARTE || '').trim();
      const wEst = String(w.ESTADO_WIP || '').trim().toUpperCase();
      if (wNP === npClean && wEst !== 'CONSUMIDO') {
        totalWipCortado += Number(w.PIEZAS_EN_WIP || 0);
      }
    });

    let totalKanbanCortado = 0;
    kanbanRows.forEach(k => {
      const kNP = String(k.NUMERO_PARTE || '').trim();
      if (kNP === npClean) {
        totalKanbanCortado += Number(k.CANTIDAD_ACTUAL || 0);
      }
    });

    const totalCortadoReal = Math.max(totalCortadoOk + totalWipCortado, totalKanbanCortado);

    let yaProgAux = 0;
    progRows.forEach(r => {
      const rNP = String(r.NUMERO_PARTE || '').trim();
      const rOp = String(r.OPERACION || '').trim().toUpperCase();
      const rEst = String(r.ESTADO_ORDEN || '').trim().toUpperCase();
      if (rNP === npClean && rOp === opUpper && rEst !== 'CANCELADA') {
        yaProgAux += Number(r.PIEZAS_PROGRAMADAS || 0);
      }
    });

    const dispAux = Math.max(0, totalCortadoReal - yaProgAux);
    const pzasProg = Number(payload.piezasProgramadas || 0);

    if (dispAux <= 0) {
      throw new Error(
        'CANDADO DE CORTE PREVIO: No se puede programar ' + payload.operacion + ' para el N/P ' + npClean +
        ' porque no hay piezas cortadas disponibles en planta (Total cortado entregado: ' + totalCortadoReal +
        ' pz | Ya programado en ' + payload.operacion + ': ' + yaProgAux + ' pz). Primero debe cortarse y entregarse en Corte Láser/Disco.'
      );
    }

    if (pzasProg > dispAux) {
      throw new Error(
        'CANDADO DE CORTE PREVIO: La cantidad programada (' + pzasProg + ' pz) excede el material cortado disponible (' +
        dispAux + ' pz disponibles para ' + payload.operacion + '). Ajuste la cantidad o espere entregas de Corte.'
      );
    }
  }

  if (totalSecTurno > 0 && rateNueva > 0 && mClean && fIso && tClean) {
    const progRows = sheetToObjects_('MES_PROGRAMACION_SEMANAL');
    let usedSec = 0;
    progRows.forEach(r => {
      if (String(r.ESTADO_ORDEN || '').trim().toUpperCase() === 'CANCELADA') return;
      if (
        String(r.MAQUINA_ASIGNADA || '').trim() === mClean &&
        normalizeIsoDate_(r.FECHA_TURNO) === fIso &&
        String(r.ID_TURNO || '').trim().toUpperCase() === tClean
      ) {
        const rSeg = Number(r.RATE_SEG || 0);
        const pMaq = Math.max(0, Number(r.PIEZAS_PROGRAMADAS || 0) - Number(r.PIEZAS_TOMADAS_DE_WIP || 0));
        if (rSeg > 0 && pMaq > 0) {
          usedSec += (pMaq * rSeg);
        }
      }
    });
    const remainingSec = Math.max(0, totalSecTurno - usedSec);

    if (isCorteOp) {
      const cortesVal = Math.max(1, Math.floor(Number(payload.cortesPorBarra || 1)));
      const time1BarSec = cortesVal * rateNueva;

      // Candado: Si queda menos tiempo del necesario para cortar 1 barra completa, el turno está bloqueado
      if (remainingSec < time1BarSec) {
        throw new Error(
          'CAPACIDAD DE TURNO AGOTADA (< 1 BARRA): En la máquina ' + mClean + ' (' + fIso + ' - ' + tClean + ') solo quedan ' +
          (Math.round((remainingSec / 60) * 10) / 10) + ' min disponibles, lo cual no alcanza para cortar 1 barra completa (' +
          (Math.round((time1BarSec / 60) * 10) / 10) + ' min requeridos para ' + cortesVal + ' pzas). Debe programar en el siguiente turno.'
        );
      }

      const maxBarrasRestantes = Math.floor(remainingSec / time1BarSec);
      const maxPzasRestantes = maxBarrasRestantes * cortesVal;
      if (pzasEnMaquinaNueva > maxPzasRestantes) {
        throw new Error(
          'CAPACIDAD DE TURNO EXCEDIDA: En la máquina ' + mClean + ' (' + fIso + ' - ' + tClean + ') quedan ' +
          (Math.round((remainingSec / 60) * 10) / 10) + ' min disponibles (máximo ' + maxBarrasRestantes +
          ' barra(s) / ' + maxPzasRestantes + ' pzas para el N/P ' + payload.numeroParte + ').'
        );
      }
    } else {
      const tiempoNecesarioSec = pzasEnMaquinaNueva * rateNueva;
      if (remainingSec < rateNueva) {
        throw new Error(
          'CAPACIDAD DE TURNO AGOTADA: En la máquina ' + mClean + ' (' + fIso + ' - ' + tClean + ') solo quedan ' +
          (Math.round((remainingSec / 60) * 10) / 10) + ' min disponibles. Debe programar en el siguiente turno.'
        );
      }
      if (tiempoNecesarioSec > remainingSec) {
        const maxPzasPorTiempo = Math.floor(remainingSec / rateNueva);
        throw new Error(
          'CAPACIDAD DE TURNO EXCEDIDA: En la máquina ' + mClean + ' (' + fIso + ' - ' + tClean + ') solo caben ' +
          maxPzasPorTiempo + ' piezas con rate de ' + rateNueva + 's (' + (Math.round((remainingSec / 60) * 10) / 10) + ' min disponibles).'
        );
      }
    }
  }

  // Si ya existe una orden activa con la misma Máquina, Fecha, Turno y N/P, acumulamos en esa misma fila para que sea una sola línea
  const progData = wsProg.getDataRange().getValues();
  const headers = progData[0];
  const idxId = headers.indexOf('ID_PROGRAMA');
  const idxFecha = headers.indexOf('FECHA_TURNO');
  const idxTurno = headers.indexOf('ID_TURNO');
  const idxMaq = headers.indexOf('MAQUINA_ASIGNADA');
  const idxNP = headers.indexOf('NUMERO_PARTE');
  const idxPzasProg = headers.indexOf('PIEZAS_PROGRAMADAS');
  const idxBarras = headers.indexOf('BARRAS_A_OCUPAR');
  const idxCortes = headers.indexOf('CORTES_POR_BARRA');
  const idxEstado = headers.indexOf('ESTADO_ORDEN');
  const idxModPor = headers.indexOf('ULTIMA_MODIFICACION_POR');
  const idxFechaMod = headers.indexOf('FECHA_ULTIMA_MODIFICACION');
  const idxColada = headers.indexOf('COMENTARIOS');

  let existingRow = -1;
  for (let i = 1; i < progData.length; i++) {
    if (
      String(progData[i][idxMaq]).trim() === mClean &&
      normalizeIsoDate_(progData[i][idxFecha]) === fIso &&
      String(progData[i][idxTurno]).trim().toUpperCase() === tClean &&
      String(progData[i][idxNP]).trim() === String(payload.numeroParte).trim() &&
      String(progData[i][idxEstado]).trim().toUpperCase() !== 'CANCELADA'
    ) {
      existingRow = i + 1;
      break;
    }
  }

  if (existingRow > 0) {
    const pzasActuales = Number(progData[existingRow - 1][idxPzasProg] || 0);
    const pzasAdicionales = Number(payload.piezasProgramadas || 0);
    const totalPzas = pzasActuales + pzasAdicionales;
    const cortesVal = Number(progData[existingRow - 1][idxCortes] || payload.cortesPorBarra || 1) || 1;
    const barrasTotal = Math.ceil(totalPzas / Math.max(1, Math.floor(cortesVal)));

    wsProg.getRange(existingRow, idxPzasProg + 1).setValue(totalPzas);
    if (idxBarras >= 0) wsProg.getRange(existingRow, idxBarras + 1).setValue(barrasTotal);
    if (payload.coladaMp && idxColada >= 0) {
      wsProg.getRange(existingRow, idxColada + 1).setValue('COLADA:' + payload.coladaMp + ' | ' + (payload.comentarios || ''));
    }
    if (idxModPor >= 0) wsProg.getRange(existingRow, idxModPor + 1).setValue(payload.usuario);
    if (idxFechaMod >= 0) wsProg.getRange(existingRow, idxFechaMod + 1).setValue(nowStr);

    logBitacora_({
      usuario: payload.usuario,
      idPrograma: progData[existingRow - 1][idxId],
      hojaDeRuta: progData[existingRow - 1][headers.indexOf('HOJA_DE_RUTA')],
      numeroParte: payload.numeroParte,
      campo: 'ACUMULACION_PIEZAS',
      valorAnterior: pzasActuales + ' pzas',
      valorNuevo: totalPzas + ' pzas (+' + pzasAdicionales + ' pzas)',
      motivo: 'Acumulación en el mismo Número de Parte, Máquina y Turno'
    });

    sendWhatsAppNotification_({
      tipoEvento: 'INCREMENTO_PROGRAMACION',
      usuarioOrigen: payload.usuario,
      destinatarios: 'TODOS',
      mensaje: '📢 [INCREMENTO PROGRAMA] Se acumularon +' + pzasAdicionales + ' pzas al N/P ' + payload.numeroParte +
               ' en ' + mClean + ' (' + fIso + ' • ' + tClean + '). Total actual: ' + totalPzas + ' pzas (' + barrasTotal + ' barras).'
    });

    return {
      ok: true,
      idPrograma: progData[existingRow - 1][idxId],
      isConsolidated: true,
      totalPiezas: totalPzas,
      message: 'Se acumularon ' + pzasAdicionales + ' piezas a la orden existente. Total: ' + totalPzas + ' piezas.'
    };
  }

  const idProg = uniqueId_('PROG');
  const comentarioColada = (payload.coladaMp ? 'COLADA:' + payload.coladaMp : 'SIN_COLADA') +
                           (payload.comentarios ? ' | ' + payload.comentarios : '');

  const newRow = [
    idProg,
    payload.semana,
    payload.fechaTurno,
    payload.idTurno,
    payload.hojaDeRuta,
    payload.cliente,
    payload.numeroParte,
    payload.operacion,
    payload.maquinaAsignada,
    payload.tipoMaquina || 'P',
    payload.rateSeg,
    payload.codigoMpUsado,
    payload.esMpAlterna || 'NO',
    payload.folioDesviacion || '',
    payload.longitudBarraMm,
    payload.cortesPorBarra,
    payload.barrasAOcupar,
    Number(payload.piezasProgramadas),
    Number(payload.piezasTomadasDeWip || 0),
    0, // PIEZAS_ENTREGADAS_OK
    0, // PIEZAS_A_WIP
    0, // PIEZAS_SCRAP
    0, // PORCENTAJE_CUMPLIMIENTO
    'PROGRAMADA',
    payload.usuario,
    nowStr,
    payload.usuario,
    nowStr,
    comentarioColada
  ];

  wsProg.appendRow(newRow);

  // Si tomó piezas de WIP disponible, descontarlas del registro de WIP
  if (payload.idWipConsumido && Number(payload.piezasTomadasDeWip) > 0) {
    consumeWipForProgram_(payload.idWipConsumido, Number(payload.piezasTomadasDeWip), payload.hojaDeRuta, payload.usuario);
  }

  sendWhatsAppNotification_({
    tipoEvento: 'NUEVA_PROGRAMACION',
    usuarioOrigen: payload.usuario,
    destinatarios: 'TODOS',
    mensaje: '📢 [NUEVO PROGRAMA MES] Semana ' + payload.semana + ' | Hoja de Ruta #' + payload.hojaDeRuta +
             ' | N/P: ' + payload.numeroParte + ' | Máq: ' + payload.maquinaAsignada +
             ' | Pzas: ' + payload.piezasProgramadas + ' | Barras: ' + payload.barrasAOcupar
  });

  return { ok: true, idPrograma: idProg, message: 'Programación detonada y notificada correctamente.' };
  });
}

/**
 * Descuenta piezas de un lote de WIP cuando Programación las asigna a una nueva corrida
 */
function consumeWipForProgram_(idWip, qtyToConsume, hojaRutaDestino, usuario) {
  const ss = getDb_();
  const wsWip = ss.getSheetByName('MES_INVENTARIO_WIP');
  const data = wsWip.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('ID_WIP');
  const idxPzas = headers.indexOf('PIEZAS_EN_WIP');
  const idxEstado = headers.indexOf('ESTADO_WIP');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idWip)) {
      const disp = Number(data[i][idxPzas] || 0);
      const restante = Math.max(0, disp - qtyToConsume);
      wsWip.getRange(i + 1, idxPzas + 1).setValue(restante);
      wsWip.getRange(i + 1, idxEstado + 1).setValue(restante === 0 ? 'CONSUMIDO EN HR ' + hojaRutaDestino : 'DISPONIBLE');
      break;
    }
  }
}

/**
 * ============================================================================
 * MÓDULO 2: SUPERVISIÓN, EMISIÓN DE ETIQUETAS FIP-55 Y SUBIDA EN PISO
 * ============================================================================
 */
function kanbanPayloadToRowObj_(payload, nowStr, idKanban, consecutivo) {
  const emptySlot = JSON.stringify({ nombre: '', emp: '', cant: '', fecha: '' });
  return {
    ID_KANBAN: idKanban,
    ID_PROGRAMA: payload.idPrograma || '',
    HOJA_DE_RUTA: payload.hojaDeRuta || '',
    NUMERO_PARTE: payload.numeroParte || '',
    CLIENTE: payload.cliente || '',
    CODIGO_MP: payload.codigoMp || '',
    LOTE_FISICO_MP: payload.coladaMp || 'PENDIENTE',
    ES_MP_ALTERNA: payload.esMpAlterna || 'NO',
    FOLIO_DESVIACION: payload.folioDesviacion || '',
    CANTIDAD_INICIAL: 0,
    CANTIDAD_ACTUAL: Number(payload.cantidadActual || 0),
    CANTIDAD_ESTANDAR_PACK: Number(payload.cantidadEstandarPack || payload.cantidadActual || 0),
    LOTE_CONTENEDOR: payload.loteContenedor || '1 de 1',
    TIPO_EMPAQUE: payload.tipoEmpaque || 'PRINCIPAL',
    NOMBRE_EMPAQUE: payload.nombreEmpaque || 'H / Rack',
    ES_PACK_MANUAL: (payload.esPackManual === true || String(payload.esPackManual || '').toUpperCase() === 'SI' || String(payload.esPackManual || '').toUpperCase() === 'TRUE') ? 'SI' : 'NO',
    CONSECUTIVO: consecutivo || payload.consecutivo || 1,
    OPERACION_ACTUAL: payload.operacionActual || '',
    MAQUINA_ACTUAL: payload.maquinaActual || '',
    ESTADO_KANBAN: payload.estadoKanban || 'EN_LLENADO',
    SIGUIENTE_OPERACION: payload.siguienteOperacion || 'INSPECCION Y LIBERACION',
    AUDITADO_FISICO: 'PENDIENTE',
    FECHA_ULTIMA_AUDITORIA: '',
    AUDITADO_POR: '',
    FECHA_CREACION: nowStr,
    ES_SEGURIDAD: payload.esSeguridad || 'NO',
    O1_JSON: emptySlot,
    O2_JSON: emptySlot,
    O3_JSON: emptySlot,
    TRAZABILIDAD_ORIGEN: payload.trazabilidadOrigen || ''
  };
}

function createKanbanCardFip55(payload) {
  // Crea un nuevo contenedor / etiqueta FIP-55 asociada a la Hoja de Ruta (alineado por encabezados reales)
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const ss = getDb_();
    const wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
    const existing = sheetToObjects_('MES_LOTES_KANBAN');

    const sameHr = existing.filter(k => String(k.HOJA_DE_RUTA) === String(payload.hojaDeRuta));
    const consecutivo = payload.consecutivo || String(sameHr.length + 1).padStart(2, '0');
    const idKanban = payload.idKanban || (payload.hojaDeRuta + '-' + consecutivo);
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    appendByHeaders_(wsKanban, kanbanPayloadToRowObj_(payload, nowStr, idKanban, consecutivo));
    return {
      ok: true,
      idKanban: idKanban,
      consecutivo: consecutivo,
      message: 'Etiqueta FIP-55 #' + idKanban + ' generada y guardada en MES_LOTES_KANBAN.'
    };
  });
}

/**
 * Emisión masiva de múltiples tarjetas FIP-55 por Standard Pack
 */
function createMultipleKanbanCardsFip55(cardsList) {
  if (!cardsList || !cardsList.length) return { ok: false, count: 0 };
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const ss = getDb_();
    const wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const lastCol = Math.max(1, wsKanban.getLastColumn());
    const headers = wsKanban.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });

    const rows = cardsList.map(function(payload, idx) {
      const obj = kanbanPayloadToRowObj_(
        payload,
        nowStr,
        payload.idKanban,
        payload.consecutivo || (idx + 1)
      );
      return rowByHeaders_(headers, obj);
    });

    if (rows.length > 0) {
      wsKanban.getRange(wsKanban.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    }
    return {
      ok: true,
      count: rows.length,
      message: rows.length + ' etiquetas FIP-55 guardadas exitosamente en MES_LOTES_KANBAN.'
    };
  });
}

/**
 * Registra aportación de operador (O1, O2, O3), valida Candado Físico de Barras,
 * persiste slots O1..O3 + trazabilidad en MES_APORTACIONES_PISO y actualiza orden/WIP.
 */
function registerFloorProductionAndUpload(payload) {
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const ss = getDb_();
    const wsProg = ss.getSheetByName('MES_PROGRAMACION_SEMANAL');
    const wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
    const wsWip = ss.getSheetByName('MES_INVENTARIO_WIP');
    const wsApt = ss.getSheetByName('MES_APORTACIONES_PISO');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const progData = wsProg.getDataRange().getValues();
    const pHeaders = progData[0].map(function(h) { return String(h).trim(); });
    const idxProgId = pHeaders.indexOf('ID_PROGRAMA');
    const idxProgHr = pHeaders.indexOf('HOJA_DE_RUTA');
    const idxProgOp = pHeaders.indexOf('OPERACION');
    const idxPzasProg = pHeaders.indexOf('PIEZAS_PROGRAMADAS');
    const idxPzasOk = pHeaders.indexOf('PIEZAS_ENTREGADAS_OK');
    const idxPzasWip = pHeaders.indexOf('PIEZAS_A_WIP');
    const idxPzasScrap = pHeaders.indexOf('PIEZAS_SCRAP');
    const idxCump = pHeaders.indexOf('PORCENTAJE_CUMPLIMIENTO');
    const idxEstadoOrd = pHeaders.indexOf('ESTADO_ORDEN');
    const idxCliente = pHeaders.indexOf('CLIENTE');
    const idxCodigoMp = pHeaders.indexOf('CODIGO_MP_USADO');
    const idxColada = pHeaders.indexOf('COMENTARIOS');

    let progRowIdx = -1;
    let pzasProgramadas = 0;
    let pzasYaEntregadas = 0;
    let pzasYaWip = 0;
    let pzasYaScrap = 0;

    // 1. Buscar coincidencia exacta por ID_PROGRAMA primero (evita cruces si distintos días/turnos comparten HR)
    const opTarget = String(payload.operacion || '').trim().toUpperCase();
    if (payload.idPrograma) {
      for (let i = 1; i < progData.length; i++) {
        if (String(progData[i][idxProgId]).trim() === String(payload.idPrograma).trim()) {
          progRowIdx = i + 1;
          break;
        }
      }
    }
    if (progRowIdx === -1 && payload.hojaDeRuta && opTarget) {
      for (let i = 1; i < progData.length; i++) {
        if (
          String(progData[i][idxProgHr]).trim() === String(payload.hojaDeRuta).trim() &&
          String(progData[i][idxProgOp] || '').trim().toUpperCase() === opTarget
        ) {
          progRowIdx = i + 1;
          break;
        }
      }
    }
    if (progRowIdx === -1 && payload.hojaDeRuta) {
      for (let i = 1; i < progData.length; i++) {
        if (String(progData[i][idxProgHr]).trim() === String(payload.hojaDeRuta).trim()) {
          progRowIdx = i + 1;
          break;
        }
      }
    }

    if (progRowIdx === -1) {
      throw new Error('No se encontró la orden de programación vinculada (HR #' + (payload.hojaDeRuta || '') + ').');
    }

    const pRow = progData[progRowIdx - 1];
    const resolvedIdProg = idxProgId >= 0 ? String(pRow[idxProgId]) : (payload.idPrograma || '');
    const resolvedProgOp = idxProgOp >= 0 ? String(pRow[idxProgOp] || '').trim().toUpperCase() : '';
    const isSameOpOrder = !opTarget || !resolvedProgOp || resolvedProgOp === opTarget;
    pzasProgramadas = Number(pRow[idxPzasProg] || 0);
    pzasYaEntregadas = Number(pRow[idxPzasOk] || 0);
    pzasYaWip = Number(pRow[idxPzasWip] || 0);
    pzasYaScrap = idxPzasScrap >= 0 ? Number(pRow[idxPzasScrap] || 0) : 0;

    const nuevasBuenas = Math.max(0, Number(payload.piezasAportadas || 0));
    const nuevasWip = Math.max(0, Number(payload.piezasExcedenteWip || 0));
    if (nuevasBuenas <= 0 && nuevasWip <= 0) {
      throw new Error('Debe ingresar una cantidad mayor a cero para registrar la subida de material.');
    }

    const totalAcumuladoPropuesto = pzasYaEntregadas + pzasYaWip + pzasYaScrap + nuevasBuenas + nuevasWip;
    const opUpper = String(payload.operacion || '').toUpperCase();
    const maqUpper = String(payload.maquina || '').toUpperCase();
    const isCorteOp = opUpper.indexOf('CORTE') !== -1 || maqUpper.indexOf('CL-') === 0 || maqUpper === 'C-123' || maqUpper === 'C-114' || maqUpper === 'C-116';

    // CANDADO FÍSICO DE MATERIA PRIMA (Solo en Corte Láser / Corte Disco donde ingresan barras)
    if (isCorteOp && isSameOpOrder && Number(payload.barrasIntroducidas) > 0) {
      const pzasEnterasPorBarra = Math.max(1, Math.floor(Number(payload.cortesPorBarra || 1)));
      const topeFisicoMaximo = Number(payload.barrasIntroducidas) * pzasEnterasPorBarra;

      if (totalAcumuladoPropuesto > topeFisicoMaximo) {
        throw new Error(
          '🚫 BLOQUEO FÍSICO DE MATERIA PRIMA: Se declararon ' + payload.barrasIntroducidas +
          ' barras introducidas (' + pzasEnterasPorBarra + ' pzas/barra = Tope físico máximo de ' +
          topeFisicoMaximo + ' piezas). El acumulado (OK + WIP + Scrap = ' + totalAcumuladoPropuesto +
          ' pzas) supera el máximo físico posible. Verifica el conteo o las barras ingresadas.'
        );
      }
    }

    // Actualizar acumulados en MES_PROGRAMACION_SEMANAL solo si la orden corresponde a esta operación
    const nuevoTotalOk = isSameOpOrder ? (pzasYaEntregadas + nuevasBuenas) : pzasYaEntregadas;
    const nuevoTotalWip = isSameOpOrder ? (pzasYaWip + nuevasWip) : pzasYaWip;
    const pctCumplimiento = pzasProgramadas > 0 ? Math.min(100, Math.round((nuevoTotalOk / pzasProgramadas) * 100)) : 100;
    const restantePorCortar = Math.max(0, pzasProgramadas - nuevoTotalOk);

    if (isSameOpOrder) {
      wsProg.getRange(progRowIdx, idxPzasOk + 1).setValue(nuevoTotalOk);
      if (idxPzasWip >= 0) wsProg.getRange(progRowIdx, idxPzasWip + 1).setValue(nuevoTotalWip);
      if (idxCump >= 0) wsProg.getRange(progRowIdx, idxCump + 1).setValue(pctCumplimiento + '%');
      if (idxEstadoOrd >= 0) wsProg.getRange(progRowIdx, idxEstadoOrd + 1).setValue(restantePorCortar === 0 ? 'COMPLETADA_EN_PISO' : 'EN_PROCESO');
    }

    // Actualizar el contenedor Kanban y el slot O1/O2/O3 en MES_LOTES_KANBAN
    const kData = wsKanban.getDataRange().getValues();
    const kHeaders = kData[0].map(function(h) { return String(h).trim(); });
    const idxKId = kHeaders.indexOf('ID_KANBAN');
    const idxCantAct = kHeaders.indexOf('CANTIDAD_ACTUAL');
    const idxCantIni = kHeaders.indexOf('CANTIDAD_INICIAL');
    const idxEstKanban = kHeaders.indexOf('ESTADO_KANBAN');
    const idxOpAct = kHeaders.indexOf('OPERACION_ACTUAL');
    const slotKey = (payload.slotOperador === 'O2' || payload.slotOperador === 'O3') ? payload.slotOperador : 'O1';
    const idxSlotCol = kHeaders.indexOf(slotKey + '_JSON');
    const idxTraza = kHeaders.indexOf('TRAZABILIDAD_ORIGEN');
    const fechaSlot = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM') + ' ' + slotKey;
    const isFinalStageOp = opUpper.indexOf('TERMINADO') !== -1 || opUpper.indexOf('INSPECC') !== -1 || opUpper === 'PT';

    let totalEnContenedor = nuevasBuenas;
    let slotObjResult = { nombre: payload.nombreOperador || '', emp: payload.numEmpleado || '', cant: nuevasBuenas, fecha: fechaSlot };
    for (let i = 1; i < kData.length; i++) {
      if (String(kData[i][idxKId]).trim() === String(payload.idKanban).trim()) {
        const prevCant = Number(kData[i][idxCantAct] || 0);
        const prevEstado = idxEstKanban >= 0 ? String(kData[i][idxEstKanban] || '').trim().toUpperCase() : '';
        // Si el contenedor viene recién avanzado desde la operación anterior (EN_PROCESO), la primera captura
        // confirma las piezas procesadas en esta etapa sin sumarlas encima del conteo de la etapa anterior.
        totalEnContenedor = (prevEstado === 'EN_PROCESO') ? nuevasBuenas : (prevCant + nuevasBuenas);
        wsKanban.getRange(i + 1, idxCantAct + 1).setValue(totalEnContenedor);
        if (idxCantIni >= 0 && prevCant === 0) {
          wsKanban.getRange(i + 1, idxCantIni + 1).setValue(totalEnContenedor);
        }
        const nuevoEstadoK = payload.cerrarContenedor
          ? (isFinalStageOp ? 'LIBERADO_PRODUCTO_TERMINADO' : 'LISTO_SIGUIENTE_PROCESO')
          : 'EN_LLENADO_PARCIAL';
        if (idxEstKanban >= 0) wsKanban.getRange(i + 1, idxEstKanban + 1).setValue(nuevoEstadoK);
        if (idxOpAct >= 0 && payload.cerrarContenedor && isFinalStageOp) {
          wsKanban.getRange(i + 1, idxOpAct + 1).setValue('PRODUCTO TERMINADO (PT)');
        }
        if (idxSlotCol >= 0) {
          const prevSlot = parseSlotJson_(kData[i][idxSlotCol]);
          const mismaOp = prevSlot.emp && payload.numEmpleado && String(prevSlot.emp) === String(payload.numEmpleado);
          const cantSlot = mismaOp ? (Number(prevSlot.cant || 0) + nuevasBuenas) : nuevasBuenas;
          slotObjResult = {
            nombre: payload.nombreOperador || prevSlot.nombre || '',
            emp: payload.numEmpleado || prevSlot.emp || '',
            cant: cantSlot,
            fecha: fechaSlot
          };
          wsKanban.getRange(i + 1, idxSlotCol + 1).setValue(JSON.stringify(slotObjResult));
        }
        if (idxTraza >= 0 && isCorteOp) {
          const prevTraza = String(kData[i][idxTraza] || '').trim();
          const trazaNueva = 'Corte: ' + (payload.nombreOperador || '-') + ' (' + (payload.numEmpleado || '-') + ') | Máq: ' + (payload.maquina || '-') + ' | ' + totalEnContenedor + ' pzas';
          wsKanban.getRange(i + 1, idxTraza + 1).setValue(prevTraza ? trazaNueva : trazaNueva);
        }
        break;
      }
    }

    // Si hubo excedente de corrida (1%-5% o lote extra para WIP), registrarlo en MES_INVENTARIO_WIP
    if (nuevasWip > 0) {
      const idWip = uniqueId_('WIP-' + (payload.hojaDeRuta || 'L'));
      appendByHeaders_(wsWip, {
        ID_WIP: idWip,
        ID_KANBAN_ORIGEN: payload.idKanban,
        HOJA_DE_RUTA: payload.hojaDeRuta,
        NUMERO_PARTE: payload.numeroParte,
        CLIENTE: payload.cliente || (idxCliente >= 0 ? pRow[idxCliente] : ''),
        OPERACION_ORIGEN: payload.operacion,
        SIGUIENTE_PROCESO: payload.siguienteOperacion || 'DOBLADO',
        MAQUINA_ORIGEN: payload.maquina,
        CODIGO_MP: payload.codigoMp || (idxCodigoMp >= 0 ? pRow[idxCodigoMp] : ''),
        PIEZAS_EN_WIP: nuevasWip,
        UBICACION_PLANTA: 'AREA WIP ' + payload.operacion,
        ESTATUS_AUDITORIA: 'PENDIENTE_AUDITORIA',
        FECHA_AUDITORIA: '',
        AUDITADO_POR: '',
        FECHA_CREACION: nowStr,
        ESTADO_WIP: 'DISPONIBLE'
      });
    }

    // Registrar genealogía en MES_APORTACIONES_PISO
    const idAportacion = uniqueId_('APT');
    if (wsApt) {
      appendByHeaders_(wsApt, {
        ID_APORTACION: idAportacion,
        FECHA_HORA: nowStr,
        ID_KANBAN: payload.idKanban,
        ID_PROGRAMA: resolvedIdProg,
        HOJA_DE_RUTA: payload.hojaDeRuta,
        NUMERO_PARTE: payload.numeroParte,
        CLIENTE: payload.cliente || (idxCliente >= 0 ? pRow[idxCliente] : ''),
        OPERACION: payload.operacion,
        MAQUINA: payload.maquina,
        SLOT: slotKey,
        NUM_EMPLEADO: payload.numEmpleado || '',
        NOMBRE_OPERADOR: payload.nombreOperador || '',
        TURNO: payload.turno || slotKey,
        PIEZAS_OK: nuevasBuenas,
        PIEZAS_WIP: nuevasWip,
        CONTENEDOR_CERRADO: payload.cerrarContenedor ? 'SI' : 'NO',
        CODIGO_MP: payload.codigoMp || (idxCodigoMp >= 0 ? pRow[idxCodigoMp] : ''),
        COLADA_MP: idxColada >= 0 ? String(pRow[idxColada] || '') : '',
        USUARIO_TERMINAL: payload.usuario || payload.nombreOperador || ''
      });
    }

    sendWhatsAppNotification_({
      tipoEvento: 'SUBIDA_MATERIAL_PISO',
      usuarioOrigen: (payload.nombreOperador || '') + ' (' + slotKey + ')',
      destinatarios: 'SUPERVISOR, SIGUIENTE_PROCESO, AUDITOR_INVENTARIO',
      mensaje: '📦 [SUBIDA DE MATERIAL] Kanban #' + payload.idKanban + ' | N/P: ' + payload.numeroParte +
               ' | Máq: ' + payload.maquina + ' | Aportación ' + slotKey + ' (' + (payload.nombreOperador || '') + '): +' + nuevasBuenas + ' pzas' +
               ' (Total Contenedor: ' + totalEnContenedor + ' pzas | Restante Orden: ' + restantePorCortar + ' pzas)' +
               (nuevasWip > 0 ? ' | +' + nuevasWip + ' pzas enviadas a WIP.' : '')
    });

    return {
      ok: true,
      idAportacion: idAportacion,
      totalEnContenedor: totalEnContenedor,
      restantePorCortar: restantePorCortar,
      nuevoTotalOk: nuevoTotalOk,
      nuevoTotalWip: nuevoTotalWip,
      pctCumplimiento: pctCumplimiento + '%',
      slotObj: slotObjResult,
      message: 'Subida registrada. Total en contenedor: ' + totalEnContenedor + ' pzas. Restante de la orden: ' + restantePorCortar + ' pzas.'
    };
  });
}

/**
 * Cierre de Turno / Segregación de Material Defectuoso en Estación (Rack Rojo Scrap / Rack Amarillo Sospechoso)
 */
function submitOperatorRackRecord(payload) {
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const ss = getDb_();
    const wsScrap = ss.getSheetByName('MES_SCRAP_REGISTROS');
    const wsProg = ss.getSheetByName('MES_PROGRAMACION_SEMANAL');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const idScrap = uniqueId_('SCR');
    const piezas = Math.max(0, Number(payload.piezasScrap || 0));
    if (piezas <= 0) {
      throw new Error('Ingrese una cantidad de piezas mayor a 0.');
    }

    const tipoRack = String(payload.tipoRack || 'ROJO').trim().toUpperCase();
    const esRojo = (tipoRack === 'ROJO');
    const estatus = esRojo ? 'SCRAP_DEFINITIVO' : 'PENDIENTE_VALIDACION';

    appendByHeaders_(wsScrap, {
      ID_SCRAP: idScrap,
      FECHA: nowStr,
      SEMANA: payload.semana || '',
      HOJA_DE_RUTA: payload.hojaDeRuta || '',
      ID_KANBAN: payload.idKanban || '',
      CLIENTE: payload.cliente || '',
      NUMERO_PARTE: payload.numeroParte || '',
      OPERACION: payload.operacion || '',
      MAQUINA: payload.maquina || '',
      TURNO: payload.turno || '',
      SUPERVISOR: payload.supervisor || '',
      OPERADOR: payload.operador || '',
      NUM_EMPLEADO: payload.numEmpleado || '',
      CODIGO_MP: payload.codigoMp || '',
      COLADA_MP: payload.coladaMp || '',
      PIEZAS_SCRAP: piezas,
      MODO_DE_FALLA: payload.modoDeFalla || (esRojo ? 'SCRAP_PISO_CIERRE_TURNO' : 'REVISION_RACK_AMARILLO'),
      DEFECTO: payload.defecto || (esRojo ? 'Defecto en proceso' : 'Sospechoso / En observación'),
      CAUSA: payload.causa || 'Reporte de Operador en Estación',
      DEPARTAMENTO: payload.departamento || 'PRODUCCION',
      AUDITADO_POR: esRojo ? 'PENDIENTE_AUDITOR' : 'EN_OBSERVACION_PISO',
      COMENTARIOS: payload.comentarios || '',
      TIPO_RACK: tipoRack,
      ESTATUS: estatus,
      ID_PROGRAMA: payload.idPrograma || '',
      CANT_RECUPERADA: 0,
      CANT_DESCARTADA: esRojo ? piezas : 0
    });

    let nuevoScrapOrden = piezas;
    if (esRojo && wsProg) {
      const pData = wsProg.getDataRange().getValues();
      const pHeaders = pData[0].map(function(h) { return String(h).trim(); });
      const idxId = pHeaders.indexOf('ID_PROGRAMA');
      const idxHR = pHeaders.indexOf('HOJA_DE_RUTA');
      const idxScrap = pHeaders.indexOf('PIEZAS_SCRAP');
      for (let i = 1; i < pData.length; i++) {
        const matchId = payload.idPrograma && idxId >= 0 && String(pData[i][idxId]).trim() === String(payload.idPrograma).trim();
        const matchHr = !payload.idPrograma && idxHR >= 0 && String(pData[i][idxHR]).trim() === String(payload.hojaDeRuta).trim();
        if (matchId || matchHr) {
          if (idxScrap >= 0) {
            nuevoScrapOrden = Number(pData[i][idxScrap] || 0) + piezas;
            wsProg.getRange(i + 1, idxScrap + 1).setValue(nuevoScrapOrden);
          }
          break;
        }
      }
    }

    sendWhatsAppNotification_({
      tipoEvento: esRojo ? 'CIERRE_TURNO_RACK_ROJO' : 'APARTADO_RACK_AMARILLO',
      usuarioOrigen: payload.operador || 'Operador',
      destinatarios: 'SUPERVISOR, CALIDAD, AUDITOR',
      mensaje: (esRojo ? '🔴 [RACK ROJO - CIERRE TURNO] ' : '🟡 [RACK AMARILLO - EN OBSERVACIÓN] ') +
               'HR #' + (payload.hojaDeRuta || '-') + ' | N/P: ' + (payload.numeroParte || '-') +
               ' | Máq: ' + (payload.maquina || '-') + ' | Cant: ' + piezas + ' pzas (' +
               (payload.defecto || tipoRack) + '). Operador: ' + (payload.operador || '-')
    });

    return {
      ok: true,
      idScrap: idScrap,
      tipoRack: tipoRack,
      estatus: estatus,
      nuevoScrapOrden: nuevoScrapOrden,
      message: esRojo
        ? '🔴 Scrap de Rack Rojo registrado (' + piezas + ' pzas) y conciliado con la orden.'
        : '🟡 ' + piezas + ' pzas apartadas a Rack Amarillo para revalidación en dispositivo.'
    };
  });
}

/**
 * Revalidación en dispositivo de control para material en Rack Amarillo (Recuperado OK vs Descartado Rojo)
 */
function resolveYellowRackRecord(payload) {
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const ss = getDb_();
    const wsScrap = ss.getSheetByName('MES_SCRAP_REGISTROS');
    const wsProg = ss.getSheetByName('MES_PROGRAMACION_SEMANAL');
    const wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
    const wsApt = ss.getSheetByName('MES_APORTACIONES_PISO');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const cantOk = Math.max(0, Number(payload.cantRecuperadaOk || 0));
    const cantRojo = Math.max(0, Number(payload.cantDescartadaRojo || 0));
    if (cantOk + cantRojo <= 0) {
      throw new Error('Especifique cuántas piezas se recuperan como Buenas o pasan a Rack Rojo.');
    }

    const sData = wsScrap.getDataRange().getValues();
    const sHeaders = sData[0].map(function(h) { return String(h).trim(); });
    const idxSId = sHeaders.indexOf('ID_SCRAP');
    const idxEst = sHeaders.indexOf('ESTATUS');
    const idxTipo = sHeaders.indexOf('TIPO_RACK');
    const idxCantRec = sHeaders.indexOf('CANT_RECUPERADA');
    const idxCantDesc = sHeaders.indexOf('CANT_DESCARTADA');
    const idxPzasScrap = sHeaders.indexOf('PIEZAS_SCRAP');
    const idxAud = sHeaders.indexOf('AUDITADO_POR');
    const idxCom = sHeaders.indexOf('COMENTARIOS');

    let targetRow = -1;
    for (let i = 1; i < sData.length; i++) {
      if (String(sData[i][idxSId]).trim() === String(payload.idScrap).trim()) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) {
      throw new Error('No se encontró el registro de Rack Amarillo #' + payload.idScrap);
    }

    const sRow = sData[targetRow - 1];
    const resolvedIdProg = payload.idPrograma || (sHeaders.indexOf('ID_PROGRAMA') >= 0 ? String(sRow[sHeaders.indexOf('ID_PROGRAMA')] || '') : '');
    const resolvedHr = payload.hojaDeRuta || (sHeaders.indexOf('HOJA_DE_RUTA') >= 0 ? String(sRow[sHeaders.indexOf('HOJA_DE_RUTA')] || '') : '');
    const resolvedKanban = payload.idKanban || (sHeaders.indexOf('ID_KANBAN') >= 0 ? String(sRow[sHeaders.indexOf('ID_KANBAN')] || '') : '');
    const resolvedNp = payload.numeroParte || (sHeaders.indexOf('NUMERO_PARTE') >= 0 ? String(sRow[sHeaders.indexOf('NUMERO_PARTE')] || '') : '');

    if (idxEst >= 0) wsScrap.getRange(targetRow, idxEst + 1).setValue('VALIDADO_DISPOSITIVO');
    if (idxTipo >= 0) wsScrap.getRange(targetRow, idxTipo + 1).setValue(cantRojo > 0 ? 'AMARILLO->ROJO' : 'AMARILLO->OK');
    if (idxCantRec >= 0) wsScrap.getRange(targetRow, idxCantRec + 1).setValue(cantOk);
    if (idxCantDesc >= 0) wsScrap.getRange(targetRow, idxCantDesc + 1).setValue(cantRojo);
    if (idxPzasScrap >= 0) wsScrap.getRange(targetRow, idxPzasScrap + 1).setValue(cantRojo);
    if (idxAud >= 0) wsScrap.getRange(targetRow, idxAud + 1).setValue(payload.usuario || 'VALIDADO_EN_PISO');
    if (idxCom >= 0) wsScrap.getRange(targetRow, idxCom + 1).setValue('Revalidado ' + nowStr + ': ' + cantOk + ' OK / ' + cantRojo + ' Rojo');

    if (wsProg) {
      const pData = wsProg.getDataRange().getValues();
      const pHeaders = pData[0].map(function(h) { return String(h).trim(); });
      const idxPId = pHeaders.indexOf('ID_PROGRAMA');
      const idxPHR = pHeaders.indexOf('HOJA_DE_RUTA');
      const idxPProg = pHeaders.indexOf('PIEZAS_PROGRAMADAS');
      const idxPOk = pHeaders.indexOf('PIEZAS_ENTREGADAS_OK');
      const idxPScrap = pHeaders.indexOf('PIEZAS_SCRAP');
      const idxPCump = pHeaders.indexOf('PORCENTAJE_CUMPLIMIENTO');
      const idxPEst = pHeaders.indexOf('ESTADO_ORDEN');
      for (let i = 1; i < pData.length; i++) {
        const mId = resolvedIdProg && idxPId >= 0 && String(pData[i][idxPId]).trim() === String(resolvedIdProg).trim();
        const mHr = !resolvedIdProg && resolvedHr && idxPHR >= 0 && String(pData[i][idxPHR]).trim() === String(resolvedHr).trim();
        if (mId || mHr) {
          const pProg = Number(pData[i][idxPProg] || 0);
          const pOkNew = Number(pData[i][idxPOk] || 0) + cantOk;
          const pScrapNew = Number(pData[i][idxPScrap] || 0) + cantRojo;
          if (cantOk > 0 && idxPOk >= 0) {
            wsProg.getRange(i + 1, idxPOk + 1).setValue(pOkNew);
            if (idxPCump >= 0) wsProg.getRange(i + 1, idxPCump + 1).setValue((pProg > 0 ? Math.min(100, Math.round((pOkNew / pProg) * 100)) : 100) + '%');
            if (idxPEst >= 0 && Math.max(0, pProg - pOkNew) === 0) wsProg.getRange(i + 1, idxPEst + 1).setValue('COMPLETADA_EN_PISO');
          }
          if (cantRojo > 0 && idxPScrap >= 0) {
            wsProg.getRange(i + 1, idxPScrap + 1).setValue(pScrapNew);
          }
          break;
        }
      }
    }

    if (cantOk > 0 && resolvedKanban && wsKanban) {
      const kData = wsKanban.getDataRange().getValues();
      const kHeaders = kData[0].map(function(h) { return String(h).trim(); });
      const idxKId = kHeaders.indexOf('ID_KANBAN');
      const idxKAct = kHeaders.indexOf('CANTIDAD_ACTUAL');
      for (let i = 1; i < kData.length; i++) {
        if (String(kData[i][idxKId]).trim() === String(resolvedKanban).trim()) {
          wsKanban.getRange(i + 1, idxKAct + 1).setValue(Number(kData[i][idxKAct] || 0) + cantOk);
          break;
        }
      }
    }

    if (cantOk > 0 && wsApt) {
      appendByHeaders_(wsApt, {
        ID_APORTACION: uniqueId_('APT-REC'),
        FECHA_HORA: nowStr,
        ID_KANBAN: resolvedKanban,
        ID_PROGRAMA: resolvedIdProg,
        HOJA_DE_RUTA: resolvedHr,
        NUMERO_PARTE: resolvedNp,
        CLIENTE: '',
        OPERACION: 'REVALIDACION RACK AMARILLO',
        MAQUINA: 'DISPOSITIVO_CONTROL',
        SLOT: 'RECUP',
        NUM_EMPLEADO: payload.numEmpleado || '',
        NOMBRE_OPERADOR: payload.usuario || 'Revalidación',
        TURNO: '',
        PIEZAS_OK: cantOk,
        PIEZAS_WIP: 0,
        CONTENEDOR_CERRADO: 'NO',
        CODIGO_MP: '',
        COLADA_MP: '',
        USUARIO_TERMINAL: payload.usuario || ''
      });
    }

    return {
      ok: true,
      idScrap: payload.idScrap,
      cantRecuperadaOk: cantOk,
      cantDescartadaRojo: cantRojo,
      message: 'Revalidación registrada: ' + cantOk + ' pzas recuperadas OK, ' + cantRojo + ' pzas descartadas a Rack Rojo.'
    };
  });
}

/**
 * Avanza una tarjeta Kanban FIP-55 a la siguiente operación / máquina en la planta
 */
function advanceKanbanStage(payload) {
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const ss = getDb_();
    const wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
    if (!wsKanban) throw new Error('Hoja MES_LOTES_KANBAN no encontrada.');

    const kData = wsKanban.getDataRange().getValues();
    const kHeaders = kData[0].map(function(h) { return String(h).trim(); });
    const idxKId = kHeaders.indexOf('ID_KANBAN');
    const idxOpAct = kHeaders.indexOf('OPERACION_ACTUAL');
    const idxMaqAct = kHeaders.indexOf('MAQUINA_ACTUAL');
    const idxEst = kHeaders.indexOf('ESTADO_KANBAN');
    const idxSigOp = kHeaders.indexOf('SIGUIENTE_OPERACION');
    const idxTraza = kHeaders.indexOf('TRAZABILIDAD_ORIGEN');
    const idxCantAct = kHeaders.indexOf('CANTIDAD_ACTUAL');
    const idxO1 = kHeaders.indexOf('O1_JSON');

    let found = false;
    for (let i = 1; i < kData.length; i++) {
      if (String(kData[i][idxKId]).trim() === String(payload.idKanban).trim()) {
        if (idxTraza !== -1) {
          const prevTraza = String(kData[i][idxTraza] || '').trim();
          const prevOp = idxOpAct !== -1 ? String(kData[i][idxOpAct] || '').trim() : '';
          const prevMaq = idxMaqAct !== -1 ? String(kData[i][idxMaqAct] || '').trim() : '';
          const prevCant = idxCantAct !== -1 ? Number(kData[i][idxCantAct] || 0) : 0;
          const o1 = idxO1 !== -1 ? parseSlotJson_(kData[i][idxO1]) : {};
          const seg = payload.trazabilidadOrigen || ((prevOp || 'Origen') + ': ' + (o1.nombre || 'Piso') + (o1.emp ? ' (' + o1.emp + ')' : '') + ' | Máq: ' + (prevMaq || '-') + ' | ' + prevCant + ' pzas');
          const alreadyHasCorte = (prevOp.toUpperCase().indexOf('CORTE') !== -1 && prevTraza.indexOf('Corte:') !== -1);
          const mergedTraza = payload.trazabilidadOrigen || (prevTraza ? ((alreadyHasCorte || prevTraza.indexOf(seg) !== -1) ? prevTraza : (prevTraza + ' → ' + seg)) : seg);
          wsKanban.getRange(i + 1, idxTraza + 1).setValue(mergedTraza);
        }
        if (idxOpAct !== -1 && payload.nuevaOperacion) wsKanban.getRange(i + 1, idxOpAct + 1).setValue(payload.nuevaOperacion);
        if (idxMaqAct !== -1 && payload.nuevaMaquina) wsKanban.getRange(i + 1, idxMaqAct + 1).setValue(payload.nuevaMaquina);
        if (idxEst !== -1 && payload.nuevoEstado) wsKanban.getRange(i + 1, idxEst + 1).setValue(payload.nuevoEstado);
        if (idxSigOp !== -1 && payload.siguienteOperacion) wsKanban.getRange(i + 1, idxSigOp + 1).setValue(payload.siguienteOperacion);
        found = true;
        break;
      }
    }

    if (!found) {
      throw new Error('No se encontró la tarjeta Kanban #' + payload.idKanban);
    }

    logBitacora_({
      usuario: payload.usuario || 'SUPERVISOR',
      idPrograma: '',
      hojaDeRuta: '',
      numeroParte: '',
      campo: 'AVANCE_ETAPA_KANBAN',
      valorAnterior: '',
      valorNuevo: payload.idKanban + ' -> ' + payload.nuevaOperacion + ' (' + (payload.nuevaMaquina || '-') + ')',
      motivo: 'Avance de etapa en piso'
    });

    return { ok: true, message: 'Contenedor #' + payload.idKanban + ' avanzado a ' + payload.nuevaOperacion };
  });
}

/**
 * ============================================================================
 * MÓDULO 3: INSPECCIÓN Y LIBERACIÓN (MUESTREO 5 PIEZAS AL FINAL DE RUTA)
 * ============================================================================
 */
function submitQualitySampling5(payload) {
  return withLock_(function() {
    const r = String(payload.rol || '').trim().toUpperCase();
    const allowed = ['CALIDAD', 'INSPECCION', 'SUPERVISOR', 'SUPERVISION', 'ADMIN', 'MASTER'];
    if (allowed.indexOf(r) === -1) {
      throw new Error('Acceso denegado: Solo Calidad / Inspección puede liberar o rechazar contenedores.');
    }

    const ss = getDb_();
    const wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
    const kData = wsKanban.getDataRange().getValues();
    const kHeaders = kData[0].map(function(h) { return String(h).trim(); });
    const idxKId = kHeaders.indexOf('ID_KANBAN');
    const idxHR = kHeaders.indexOf('HOJA_DE_RUTA');
    const idxEst = kHeaders.indexOf('ESTADO_KANBAN');
    const idxOpAct = kHeaders.indexOf('OPERACION_ACTUAL');

    const malasEnMuestra = Number(payload.piezasFueraSpec || 0);

    // ESCENARIO A: 5 de 5 piezas OK -> Se libera a PRODUCTO TERMINADO (PT)
    if (malasEnMuestra === 0) {
      const idxCantAct = kHeaders.indexOf('CANTIDAD_ACTUAL');
      const idxCantIni = kHeaders.indexOf('CANTIDAD_INICIAL');
      const pzasLib = Number(payload.piezasLiberadas || 0);
      for (let i = 1; i < kData.length; i++) {
        if (String(kData[i][idxKId]).trim() === String(payload.idKanban).trim()) {
          if (pzasLib > 0 && idxCantAct >= 0) {
            wsKanban.getRange(i + 1, idxCantAct + 1).setValue(pzasLib);
            if (idxCantIni >= 0 && Number(kData[i][idxCantIni] || 0) === 0) {
              wsKanban.getRange(i + 1, idxCantIni + 1).setValue(pzasLib);
            }
          }
          wsKanban.getRange(i + 1, idxEst + 1).setValue('LIBERADO_PRODUCTO_TERMINADO');
          wsKanban.getRange(i + 1, idxOpAct + 1).setValue('PRODUCTO TERMINADO (PT)');
          break;
        }
      }

      sendWhatsAppNotification_({
        tipoEvento: 'LIBERACION_PT',
        usuarioOrigen: payload.inspector,
        destinatarios: 'SUPERVISOR, ALMACEN_PT',
        mensaje: '🟢 [LIBERADO A PT] Contenedor Kanban #' + payload.idKanban + ' (N/P ' + payload.numeroParte +
                 ', HR #' + payload.hojaDeRuta + ') aprobó muestreo 5/5 OK. Pasa a Producto Terminado.'
      });

      return {
        ok: true,
        status: 'LIBERADO_PT',
        message: 'Muestreo 5/5 APROBADO. Contenedor liberado a Producto Terminado (Etiqueta Verde).'
      };
    }

    // ESCENARIO B: >= 1 pieza de las 5 sale FUERA DE ESPECIFICACIÓN
    // Regla: Se rechazan TODOS los contenedores de esa Hoja de Ruta y se para proceso para inspección al 100%
    let contenedoresBloqueados = 0;
    for (let i = 1; i < kData.length; i++) {
      if (String(kData[i][idxHR]).trim() === String(payload.hojaDeRuta).trim()) {
        wsKanban.getRange(i + 1, idxEst + 1).setValue('CUARENTENA_CONTENCION_100%');
        contenedoresBloqueados++;
      }
    }

    // Marcar la orden en MES_PROGRAMACION_SEMANAL como PARO_POR_CALIDAD
    const wsProg = ss.getSheetByName('MES_PROGRAMACION_SEMANAL');
    const pData = wsProg.getDataRange().getValues();
    const pHeaders = pData[0].map(function(h) { return String(h).trim(); });
    const idxPHR = pHeaders.indexOf('HOJA_DE_RUTA');
    const idxPEst = pHeaders.indexOf('ESTADO_ORDEN');
    for (let i = 1; i < pData.length; i++) {
      if (String(pData[i][idxPHR]).trim() === String(payload.hojaDeRuta).trim()) {
        wsProg.getRange(i + 1, idxPEst + 1).setValue('🚨 PARO_POR_CALIDAD_INSPECCION_100%');
      }
    }

    sendWhatsAppNotification_({
      tipoEvento: 'ALERTA_CONTENCION_CALIDAD',
      usuarioOrigen: payload.inspector,
      destinatarios: 'TODOS',
      mensaje: '🚨 [PARAR PROCESO - RECHAZO EN MUESTREO 5 PZAS] N/P: ' + payload.numeroParte +
               ' | Hoja de Ruta #' + payload.hojaDeRuta + ' | Contenedor #' + payload.idKanban +
               ' presentó ' + malasEnMuestra + '/5 pzas fuera de especificación (' + (payload.defectoPreliminar || 'Defecto') + '). ' +
               'Se bloquearon en CUARENTENA los ' + contenedoresBloqueados + ' contenedores de la corrida para INSPECCIÓN AL 100%.'
    });

    return {
      ok: true,
      status: 'CUARENTENA_TOTAL',
      contenedoresBloqueados: contenedoresBloqueados,
      message: '🚨 CONTENCIÓN ACTIVADA: Se pusieron en CUARENTENA los ' + contenedoresBloqueados +
               ' contenedores de la Hoja de Ruta #' + payload.hojaDeRuta + ' y se envió alerta de PARO DE PROCESO.'
    };
  });
}

/**
 * Resolución de Cuarentena tras Inspección al 100%
 * Separa piezas en: Producto Terminado, Retrabajo Interno (ocultable) y Scrap Definitivo
 */
function resolveQuarantine100(payload) {
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const ss = getDb_();
    const wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
    const kData = wsKanban.getDataRange().getValues();
    const kHeaders = kData[0].map(function(h) { return String(h).trim(); });
    const idxKId = kHeaders.indexOf('ID_KANBAN');
    const idxHR = kHeaders.indexOf('HOJA_DE_RUTA');
    const idxCantAct = kHeaders.indexOf('CANTIDAD_ACTUAL');
    const idxEst = kHeaders.indexOf('ESTADO_KANBAN');
    const idxOpAct = kHeaders.indexOf('OPERACION_ACTUAL');

    let hojaRuta = payload.hojaDeRuta || '';
    for (let i = 1; i < kData.length; i++) {
      if (String(kData[i][idxKId]).trim() === String(payload.idKanban).trim()) {
        if (!hojaRuta && idxHR >= 0) hojaRuta = String(kData[i][idxHR] || '');
        wsKanban.getRange(i + 1, idxCantAct + 1).setValue(Number(payload.piezasLiberadasPt || 0));
        if (Number(payload.piezasRetrabajoInterno || 0) > 0) {
          wsKanban.getRange(i + 1, idxEst + 1).setValue('EN_RETRABAJO_INTERNO (' + payload.piezasRetrabajoInterno + ' pzas)');
        } else {
          wsKanban.getRange(i + 1, idxEst + 1).setValue('LIBERADO_PRODUCTO_TERMINADO');
          wsKanban.getRange(i + 1, idxOpAct + 1).setValue('PRODUCTO TERMINADO (PT)');
        }
        break;
      }
    }

    const pzasScrap = Math.max(0, Number(payload.piezasScrap || 0));
    if (pzasScrap > 0 && hojaRuta) {
      const wsProg = ss.getSheetByName('MES_PROGRAMACION_SEMANAL');
      const pData = wsProg.getDataRange().getValues();
      const pHeaders = pData[0].map(function(h) { return String(h).trim(); });
      const idxPHR = pHeaders.indexOf('HOJA_DE_RUTA');
      const idxPScrap = pHeaders.indexOf('PIEZAS_SCRAP');
      for (let i = 1; i < pData.length; i++) {
        if (String(pData[i][idxPHR]).trim() === String(hojaRuta).trim() && idxPScrap >= 0) {
          wsProg.getRange(i + 1, idxPScrap + 1).setValue(Number(pData[i][idxPScrap] || 0) + pzasScrap);
          break;
        }
      }
    }

    return {
      ok: true,
      message: 'Inspección 100% registrada: ' + payload.piezasLiberadasPt + ' pzas a PT, ' +
               payload.piezasRetrabajoInterno + ' pzas a Retrabajo Interno, ' +
               pzasScrap + ' pzas turnadas al Auditor de Scrap.'
    };
  });
}

/**
 * ============================================================================
 * MÓDULO 4: AUDITOR DE INVENTARIO FÍSICO Y SCRAP (MISMA PERSONA)
 * ============================================================================
 */
function submitPhysicalInventoryAudit(payload) {
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const r = String(payload.rol || '').trim().toUpperCase();
    if (r !== 'AUDITOR_INVENTARIO' && r !== 'AUDITORIA' && r !== 'ADMIN' && r !== 'MASTER') {
      throw new Error('Acceso denegado: Solo el Auditor de Inventario o Master puede registrar conciliaciones físicas.');
    }

    const ss = getDb_();
    const wsAudit = ss.getSheetByName('MES_AUDITORIA_FISICA');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const cantSis = Number(payload.cantidadSistema || 0);
    const cantFis = Number(payload.cantidadFisica || 0);
    const dif = cantFis - cantSis;
    const resultado = dif === 0 ? 'CONCILIADO_OK' : 'AJUSTE_POR_DIFERENCIA';
    const idAudit = uniqueId_('AUD');

    wsAudit.appendRow([
      idAudit,
      nowStr,
      payload.usuario,
      payload.tipoInventario, // 'WIP' o 'KANBAN_PT'
      payload.idRegistro,
      payload.numeroParte,
      cantSis,
      cantFis,
      dif,
      resultado,
      payload.motivoAjuste || 'Verificación física en piso'
    ]);

    // Actualizar sello de auditoría en la tabla correspondiente
    if (payload.tipoInventario === 'WIP') {
      const wsWip = ss.getSheetByName('MES_INVENTARIO_WIP');
      const wData = wsWip.getDataRange().getValues();
      const wHeaders = wData[0].map(function(h) { return String(h).trim(); });
      const idxId = wHeaders.indexOf('ID_WIP');
      const idxPzas = wHeaders.indexOf('PIEZAS_EN_WIP');
      const idxEstAud = wHeaders.indexOf('ESTATUS_AUDITORIA');
      const idxAudPor = wHeaders.indexOf('AUDITADO_POR');
      const idxFecAud = wHeaders.indexOf('FECHA_AUDITORIA');

      for (let i = 1; i < wData.length; i++) {
        if (String(wData[i][idxId]).trim() === String(payload.idRegistro).trim()) {
          wsWip.getRange(i + 1, idxPzas + 1).setValue(cantFis);
          wsWip.getRange(i + 1, idxEstAud + 1).setValue('VERIFICADO_FISICO');
          wsWip.getRange(i + 1, idxAudPor + 1).setValue(payload.usuario);
          wsWip.getRange(i + 1, idxFecAud + 1).setValue(nowStr);
          break;
        }
      }
    } else {
      const wsKanban = ss.getSheetByName('MES_LOTES_KANBAN');
      const kData = wsKanban.getDataRange().getValues();
      const kHeaders = kData[0].map(function(h) { return String(h).trim(); });
      const idxId = kHeaders.indexOf('ID_KANBAN');
      const idxCant = kHeaders.indexOf('CANTIDAD_ACTUAL');
      const idxAud = kHeaders.indexOf('AUDITADO_FISICO');
      const idxFec = kHeaders.indexOf('FECHA_ULTIMA_AUDITORIA');
      const idxPor = kHeaders.indexOf('AUDITADO_POR');

      for (let i = 1; i < kData.length; i++) {
        if (String(kData[i][idxId]).trim() === String(payload.idRegistro).trim()) {
          wsKanban.getRange(i + 1, idxCant + 1).setValue(cantFis);
          wsKanban.getRange(i + 1, idxAud + 1).setValue('VERIFICADO_FISICO');
          wsKanban.getRange(i + 1, idxFec + 1).setValue(nowStr);
          wsKanban.getRange(i + 1, idxPor + 1).setValue(payload.usuario);
          break;
        }
      }
    }

    if (dif !== 0) {
      sendWhatsAppNotification_({
        tipoEvento: 'DIFERENCIA_AUDITORIA_FISICA',
        usuarioOrigen: payload.usuario,
        destinatarios: 'PROGRAMACION, SUPERVISOR, ADMIN',
        mensaje: '🔍 [AJUSTE DE AUDITORÍA FÍSICA] ' + payload.tipoInventario + ' #' + payload.idRegistro +
                 ' (N/P ' + payload.numeroParte + '). Sistema decía: ' + cantSis +
                 ' pzas | Físico real: ' + cantFis + ' pzas (Diferencia: ' + (dif > 0 ? '+' : '') + dif +
                 ' pzas). Motivo: ' + (payload.motivoAjuste || 'Ajuste en conteo físico')
      });
    }

    return { ok: true, idAudit: idAudit, resultado: resultado, diferencia: dif, message: 'Auditoría física registrada correctamente.' };
  });
}

/**
 * Registro rápido de Scrap por el Auditor (16 campos se llenan solos desde la orden/QR)
 */
function submitScrapAuditRecord(payload) {
  return withLock_(function() {
    ensureDatabaseInitialized_();
    const ss = getDb_();
    const wsScrap = ss.getSheetByName('MES_SCRAP_REGISTROS');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const idScrap = uniqueId_('SCR');
    const pzasScrap = Number(payload.piezasScrap || 0);

    appendByHeaders_(wsScrap, {
      ID_SCRAP: idScrap,
      FECHA: nowStr,
      SEMANA: payload.semana || '',
      HOJA_DE_RUTA: payload.hojaDeRuta || '',
      ID_KANBAN: payload.idKanban || '',
      CLIENTE: payload.cliente || '',
      NUMERO_PARTE: payload.numeroParte || '',
      OPERACION: payload.operacion || '',
      MAQUINA: payload.maquina || '',
      TURNO: payload.turno || '',
      SUPERVISOR: payload.supervisor || '',
      OPERADOR: payload.operador || '',
      CODIGO_MP: payload.codigoMp || '',
      COLADA_MP: payload.coladaMp || '',
      PIEZAS_SCRAP: pzasScrap,
      MODO_DE_FALLA: payload.modoDeFalla || '',
      DEFECTO: payload.defecto || '',
      CAUSA: payload.causa || '',
      DEPARTAMENTO: payload.departamento || 'PRODUCCION',
      AUDITADO_POR: payload.usuarioAuditor || '',
      COMENTARIOS: payload.comentarios || '',
      TIPO_RACK: payload.tipoRack || 'ROJO',
      ESTATUS: 'AUDITADO',
      ID_PROGRAMA: payload.idPrograma || '',
      CANT_RECUPERADA: 0,
      CANT_DESCARTADA: pzasScrap
    });

    // Sumar al acumulado de Scrap de la orden en MES_PROGRAMACION_SEMANAL
    const wsProg = ss.getSheetByName('MES_PROGRAMACION_SEMANAL');
    const pData = wsProg.getDataRange().getValues();
    const pHeaders = pData[0].map(function(h) { return String(h).trim(); });
    const idxId = pHeaders.indexOf('ID_PROGRAMA');
    const idxHR = pHeaders.indexOf('HOJA_DE_RUTA');
    const idxScrap = pHeaders.indexOf('PIEZAS_SCRAP');
    for (let i = 1; i < pData.length; i++) {
      const mId = payload.idPrograma && idxId >= 0 && String(pData[i][idxId]).trim() === String(payload.idPrograma).trim();
      const mHr = !payload.idPrograma && idxHR >= 0 && String(pData[i][idxHR]).trim() === String(payload.hojaDeRuta).trim();
      if (mId || mHr) {
        const actual = Number(pData[i][idxScrap] || 0);
        wsProg.getRange(i + 1, idxScrap + 1).setValue(actual + pzasScrap);
        break;
      }
    }

    return { ok: true, idScrap: idScrap, message: 'Scrap auditado y conciliado correctamente.' };
  });
}

/**
 * ============================================================================
 * MÓDULO 5: ADMIN, SWITCH DE MODO AUDITORÍA EXTERNA, WHATSAPP Y CORTE ANUAL
 * ============================================================================
 */
function toggleExternalAuditMode(enable, usuario) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('MODO_AUDITORIA_EXTERNA', enable ? 'TRUE' : 'FALSE');
  return {
    ok: true,
    modoAuditoriaExterna: Boolean(enable),
    message: enable
      ? '🔒 MODO AUDITORÍA EXTERNA ACTIVADO: Se han ocultado todas las vistas y registros de retrabajos internos.'
      : '🔓 Modo Operación Interna restaurado: Control de retrabajos visible.'
  };
}

/**
 * Ejecuta el Corte / Archivado Anual moviendo únicamente órdenes y lotes CERRADOS/TERMINADOS
 * y conservando todos los Catálogos y el WIP disponible vivo.
 */
function executeAnnualCutoffArchive(yearLabel, usuario) {
  return withLock_(function() {
    const ss = getDb_();
    const archiveTitle = 'MES_TEFLU_HISTORICO_' + yearLabel;
    const archiveSs = SpreadsheetApp.create(archiveTitle);
    const defaultSheet = archiveSs.getSheets()[0];

    const tablesToArchive = [
      { name: 'MES_PROGRAMACION_SEMANAL', statusCol: 'ESTADO_ORDEN', closedValues: ['COMPLETADA_EN_PISO', 'CERRADA', 'TERMINADA'] },
      { name: 'MES_LOTES_KANBAN', statusCol: 'ESTADO_KANBAN', closedValues: ['LIBERADO_PRODUCTO_TERMINADO', 'CERRADO'] },
      { name: 'MES_SCRAP_REGISTROS', statusCol: 'ESTATUS', closedValues: ['AUDITADO', 'VALIDADO_DISPOSITIVO', 'SCRAP_DEFINITIVO'] },
      { name: 'MES_APORTACIONES_PISO', statusCol: null, closedValues: [] },
      { name: 'MES_BITACORA_CAMBIOS', statusCol: null, closedValues: [] },
      { name: 'MES_AUDITORIA_FISICA', statusCol: null, closedValues: [] },
      { name: 'MES_NOTIFICACIONES_LOG', statusCol: null, closedValues: [] }
    ];

    let totalArchived = 0;

    tablesToArchive.forEach(cfg => {
      const ws = ss.getSheetByName(cfg.name);
      if (!ws) return;
      const data = ws.getDataRange().getValues();
      if (data.length < 2) return;

      const headers = data[0];
      const idxStatus = cfg.statusCol ? headers.indexOf(cfg.statusCol) : -1;
      const rowsForArchive = [headers];
      const rowsToKeep = [headers];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (idxStatus === -1 || cfg.closedValues.indexOf(String(row[idxStatus]).trim()) !== -1) {
          rowsForArchive.push(row);
          totalArchived++;
        } else {
          rowsToKeep.push(row);
        }
      }

      const archWs = archiveSs.insertSheet(cfg.name);
      archWs.getRange(1, 1, rowsForArchive.length, headers.length).setValues(rowsForArchive);

      ws.clearContents();
      ws.getRange(1, 1, rowsToKeep.length, headers.length).setValues(rowsToKeep);
    });

    if (archiveSs.getSheets().length > 1 && defaultSheet) {
      archiveSs.deleteSheet(defaultSheet);
    }

    return {
      ok: true,
      archiveUrl: archiveSs.getUrl(),
      totalArchived: totalArchived,
      message: 'Corte Anual ' + yearLabel + ' completado. Se archivaron ' + totalArchived +
               ' registros cerrados y se conservaron intactos los Catálogos y el WIP activo.'
    };
  });
}

/**
 * Registra en MES_BITACORA_CAMBIOS
 */
function logBitacora_(entry) {
  const ss = getDb_();
  const ws = ss.getSheetByName('MES_BITACORA_CAMBIOS');
  if (!ws) return;
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const idCambio = 'CHG-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  ws.appendRow([
    idCambio,
    nowStr,
    entry.usuario,
    entry.idPrograma || '',
    entry.hojaDeRuta || '',
    entry.numeroParte || '',
    entry.campo || '',
    entry.valorAnterior || '',
    entry.valorNuevo || '',
    entry.motivo || ''
  ]);
}

/**
 * Registra la notificación en MES_NOTIFICACIONES_LOG y la envía por WhatsApp si hay Webhook/API configurado
 */
function sendWhatsAppNotification_(notif) {
  const ss = getDb_();
  const ws = ss.getSheetByName('MES_NOTIFICACIONES_LOG');
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const idNotif = 'NTF-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');

  let estatus = 'REGISTRADA_EN_PANEL';
  const props = PropertiesService.getScriptProperties();
  const webhookUrl = props.getProperty('WHATSAPP_WEBHOOK_URL');

  if (webhookUrl) {
    try {
      UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          event: notif.tipoEvento,
          from: notif.usuarioOrigen,
          to: notif.destinatarios,
          message: notif.mensaje
        }),
        muteHttpExceptions: true
      });
      estatus = 'ENVIADA_WHATSAPP';
    } catch (e) {
      estatus = 'ERROR_API: ' + e.message;
    }
  }

  if (ws) {
    ws.appendRow([
      idNotif,
      nowStr,
      notif.tipoEvento,
      notif.usuarioOrigen,
      notif.destinatarios,
      'WHATSAPP / PANEL MES',
      notif.mensaje,
      estatus
    ]);
  }
}
