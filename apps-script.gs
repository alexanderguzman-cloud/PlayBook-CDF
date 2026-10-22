/**
 * SomosCX · Registro de clics (una fila por clic), comentarios y preguntas sin respuesta del asistente
 * Pega este código en: Sheet «SomosCX · Clics por usuario» › Extensiones › Apps Script.
 * 1) Configuración del proyecto › Propiedades del script: agrega TOKEN con una clave larga (la misma que SHEETS_TOKEN en Netlify).
 * 2) Ejecuta una vez la función «preparar» (autoriza los permisos).
 * 3) Implementar › Nueva implementación › Aplicación web · Ejecutar como: Yo · Acceso: Cualquier persona.
 *    Copia la URL que termina en /exec y pégala en Netlify como SHEETS_URL.
 */
const CLICS_ID = '1BlKQMdubC2RWKFw-l7YwPjf5X4SFyLIJ8SFxndm5Pxs';
const FEEDBACK_ID = '1IV6wHn1VHYwTHVhvOLCIRCHeK4avvCPCK1X1-J2SI08';
const TZ = 'America/Bogota';

function doPost(e) {
  let b = {};
  try { b = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return salida({ ok: false }); }
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || b.token !== token) return salida({ ok: false, error: 'token' });
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    if (b.tipo === 'clics') registrarClics(b);
    else if (b.tipo === 'feedback') registrarFeedback(b);
    else if (b.tipo === 'asistente') registrarAsistente(b);
    return salida({ ok: true });
  } finally { lock.releaseLock(); }
}

function hojaClics() {
  const c = SpreadsheetApp.openById(CLICS_ID);
  return c.getSheetByName('Clics por día') || c.getSheets()[0];
}

/** Nombre corto de cada proceso (el mismo que muestra SomosCX). */
const CORTO = {"TRA-01": "Buscar al cliente", "TRA-02": "Validar identidad", "TRA-03": "Planes y equipos", "TRA-04": "Escalamientos", "TRA-05": "PQRS", "FAC-01": "Factura", "FAC-02": "Pagos", "FAC-03": "Factura electrónica", "FAC-04": "Prorrateo", "FAC-05": "Facturación B2B", "FAC-06": "Descuentos", "FAC-07": "Cupones", "FAC-08": "Referidos", "FAC-09": "Tokenización", "FAC-10": "My Somos", "FAC-11": "Link de pago", "FAC-12": "Pagos adelantados", "FAC-13": "Suspendido por pago", "FAC-14": "Reembolsos", "PET-01": "Traslados", "PET-02": "Agendamientos", "PET-03": "Cambio de plan", "PET-04": "Servicios adicionales", "PET-05": "Somos TV", "PET-06": "Cambio de titular", "PET-07": "Suspensión temporal", "PET-08": "Cancelación", "PET-09": "Devolución de equipos", "PET-10": "Actualizar datos", "PET-11": "Ventas nuevas", "CER-01": "Certificados", "GUI-01": "Guía: Referidos", "GUI-02": "Guía: Tokenización", "GUI-03": "Guía: My Somos", "GUI-04": "Guía: Factura", "MAS-01": "¿Está en un masivo?", "MAS-02": "Ticket de masivo", "MAS-03": "Diagnóstico por hipótesis", "MAS-04": "Masivo en Grafana", "MAS-05": "Topologías del SITE", "CFG-01": "IP fija pública", "CFG-02": "Abrir puertos", "CFG-03": "Modo Bridge", "CFG-04": "DMZ", "CFG-05": "Direccionamiento IP", "ESP-01": "Reset de fábrica", "ESP-02": "Verificar en el Edge", "ESP-03": "LAN y WiFi", "ESP-04": "IP fija pública", "ESP-05": "Abrir puertos", "STV-01": "HDMI sin señal", "STV-02": "App desactualizada", "STV-03": "TV Box sin internet", "STV-04": "Iniciar sesión (QR)", "STV-05": "Fallas en canales", "STV-06": "Flecha en pantalla", "STV-07": "Errores raros (MAC)", "STV-08": "Control dañado", "STV-09": "Control sin luz", "STV-10": "Control no prende el TV Box", "STV-11": "Errores en pantalla", "GRA-01": "Cliente en Grafana", "GRA-02": "SITE en Grafana", "PHA-01": "Phantom offline", "PHA-02": "Intermitencia Phantom", "PHA-03": "Phantom conectado sin navegar", "PHA-04": "Bajas velocidades Phantom", "PHA-05": "Cobertura WiFi Phantom", "PHA-06": "Cambio de WiFi Phantom", "EER-01": "Eero offline", "EER-02": "Intermitencia Eero", "EER-03": "Eero conectado sin navegar", "EER-04": "Bajas velocidades Eero", "EER-05": "Cobertura WiFi Eero", "EER-06": "Cambio de WiFi Eero", "DEC-01": "Deco offline", "DEC-02": "Intermitencia Deco", "DEC-03": "Deco conectado sin navegar", "DEC-04": "Bajas velocidades Deco", "DEC-05": "Cobertura WiFi Deco", "DEC-06": "Cambio de WiFi Deco", "VIL-01": "Vilo offline", "VIL-02": "Intermitencia Vilo", "VIL-03": "Vilo conectado sin navegar", "VIL-04": "Bajas velocidades Vilo", "VIL-05": "Cobertura WiFi Vilo", "VIL-06": "Cambio de WiFi Vilo", "APP-01": "Streaming", "APP-02": "Páginas web", "APP-03": "VPN", "FAC-15": "Pre-invoice"};

/** Proceso principal a partir del texto registrado: «PET-01 · Traslados › Paso 01 · …» o «PET-01, FAC-02» → «Traslados». */
function procesoPrincipal(txt) {
  const L = [], re = /\b([A-Z]{2,4}-\d{2})\b(?:\s*·\s*([^›,|]+?)\s*(?=›|,|\||$))?/g; let m;
  while ((m = re.exec(String(txt || '')))) { const n = (CORTO[m[1]] || (m[2] || '').trim()); if (n && L.indexOf(n) < 0) L.push(n); }
  return L.join(', ');
}

/** Una fila por cada clic: Fecha | Correo | Nombre | Clics | Procesos abiertos | Última actividad | Búsquedas | Módulo | Acción | Proceso */
function registrarClics(b) {
  const hoja = hojaClics();
  if (!hoja.getRange('G1').getValue()) hoja.getRange('G1').setValue('Búsquedas');
  if (!hoja.getRange('H1').getValue()) hoja.getRange('H1:I1').setValues([['Módulo', 'Acción']]);
  if (!hoja.getRange('J1').getValue()) hoja.getRange('J1').setValue('Proceso').setFontWeight('bold').setBackground('#0A0A0A').setFontColor('#FFFFFF');
  const email = String(b.email || '').toLowerCase(), nombre = b.nombre || '';
  const ev = Array.isArray(b.eventos) ? b.eventos : [];
  const filas = ev.map(e => {
    const f = new Date(Number(e.ts) || Date.now());
    const proc = [String(e.proceso || ''), String(e.paso || '')].filter(String).join(' › ');
    return [Utilities.formatDate(f, TZ, 'yyyy-MM-dd'), email, nombre, 1, proc, Utilities.formatDate(f, TZ, 'yyyy-MM-dd HH:mm:ss'),
      String(e.busqueda || ''), String(e.modulo || ''), String(e.accion || ''), procesoPrincipal(proc) || procesoPrincipal(e.accion)];
  });
  // compatibilidad con la versión anterior de la página (envío agrupado)
  if (!filas.length && (Number(b.clics) || (b.procesos || []).length)) {
    const ahora = new Date();
    filas.push([Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd'), email, nombre, Number(b.clics) || 0, (b.procesos || []).join(', '),
      Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd HH:mm:ss'), (b.busquedas || []).join(' | '), '', '', procesoPrincipal((b.procesos || []).join(', '))]);
  }
  if (!filas.length) return;
  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 10).setValues(filas);
}

function registrarFeedback(b) {
  const f = SpreadsheetApp.openById(FEEDBACK_ID), hoja = f.getSheetByName('Comentarios') || f.getSheets()[0];
  hoja.appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), String(b.email || '').toLowerCase(), b.proceso || '', b.paso || '', b.clase || '', b.comentario || '', 'Pendiente', '', b.pagina || '']);
}

/** Ejecútala una vez: llena la columna «Proceso» (J) en las filas que ya estaban cargadas. */
function rellenarProceso() {
  const hoja = hojaClics(), n = hoja.getLastRow();
  hoja.getRange('J1').setValue('Proceso').setFontWeight('bold').setBackground('#0A0A0A').setFontColor('#FFFFFF');
  if (n < 2) return;
  const d = hoja.getRange(2, 5, n - 1, 5).getValues(); // E (procesos) … I (acción)
  const out = d.map(r => [procesoPrincipal(r[0]) || procesoPrincipal(r[4])]);
  hoja.getRange(2, 10, n - 1, 1).setValues(out);
}

const ASIS = 'Preguntas sin respuesta';
function hojaAsistente() {
  const f = SpreadsheetApp.openById(FEEDBACK_ID); let h = f.getSheetByName(ASIS);
  if (!h) {
    h = f.insertSheet(ASIS, f.getSheets().length);
    h.appendRow(['Fecha', 'Analista', 'Nombre', 'Pregunta al asistente', 'Dónde estaba', 'Estado', 'Proceso donde se agregó', 'Notas']);
    h.setFrozenRows(1); h.getRange('A1:H1').setFontWeight('bold').setBackground('#0A0A0A').setFontColor('#FFFFFF');
    h.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pendiente', 'Agregada', 'No aplica'], true).build());
    h.setColumnWidth(4, 420);
  }
  return h;
}

function registrarAsistente(b) {
  hojaAsistente().appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), String(b.email || '').toLowerCase(), b.nombre || '', b.pregunta || '', b.vista || '', 'Pendiente', '', '']);
}

/** Lo más visto y lo más buscado de los últimos 30 días (lo usa la página para «Búsquedas frecuentes»). */
function doGet(e) {
  const p = (e && e.parameter) || {}, token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || p.token !== token) return salida({ ok: false, error: 'token' });
  if (p.tipo !== 'top') return salida({ ok: true });
  const datos = hojaClics().getDataRange().getValues(), desde = new Date(Date.now() - 30 * 864e5);
  const cp = {}, cb = {};
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i][0] instanceof Date ? datos[i][0] : new Date(String(datos[i][0]));
    if (isNaN(f) || f < desde) continue;
    String(datos[i][4] || '').split(',').map(x => x.trim().split(' ')[0]).filter(x => /^[A-Z]{3}-\d{2}$/.test(x)).forEach(x => cp[x] = (cp[x] || 0) + 1);
    String(datos[i][6] || '').split(' | ').map(x => x.trim()).filter(x => x.length >= 3).forEach(x => cb[x] = (cb[x] || 0) + 1);
  }
  const top = (o, n) => Object.keys(o).sort((a, b) => o[b] - o[a]).slice(0, n);
  return salida({ ok: true, procesos: top(cp, 8), busquedas: top(cb, 10) });
}

function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Ejecútala una vez: da formato a los dos Sheets y crea el resumen por usuario. */
function preparar() {
  const c = SpreadsheetApp.openById(CLICS_ID), hc = c.getSheets()[0];
  hc.setName('Clics por día'); hc.setFrozenRows(1); hc.getRange('G1:I1').setValues([[hc.getRange('G1').getValue() || 'Búsquedas', 'Módulo', 'Acción']]); hc.getRange('A1:I1').setFontWeight('bold').setBackground('#0A0A0A').setFontColor('#FFFFFF');
  let r = c.getSheetByName('Resumen por usuario'); if (!r) r = c.insertSheet('Resumen por usuario');
  r.clear();
  r.getRange('A1').setFormula("=QUERY('Clics por día'!A:F, \"select B, sum(D), max(F) where B is not null group by B order by sum(D) desc label B 'Correo', sum(D) 'Clics totales', max(F) 'Última actividad'\", 1)");
  r.getRange('D1').setValue('Días activos');
  r.getRange('D2').setFormula("=MAP(A2:A, LAMBDA(c, IF(c=\"\", \"\", COUNTUNIQUEIFS('Clics por día'!A:A, 'Clics por día'!B:B, c))))");
  r.getRange('A1:D1').setFontWeight('bold');
  const f = SpreadsheetApp.openById(FEEDBACK_ID), hf = f.getSheets()[0];
  hf.setName('Comentarios'); hf.setFrozenRows(1); hf.getRange('A1:I1').setFontWeight('bold').setBackground('#0A0A0A').setFontColor('#FFFFFF');
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(['Pendiente', 'En revisión', 'Resuelto', 'No aplica'], true).build();
  hf.getRange('G2:G').setDataValidation(regla);
  hojaAsistente();
}
