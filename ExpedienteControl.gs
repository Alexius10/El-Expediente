/**
 * EL EXPEDIENTE · Control de acceso, validación y sorteo  (versión de un solo día)
 * ------------------------------------------------------------------------------
 * Script de Google Apps Script vinculado a la hoja de cálculo del juego.
 *
 * Cómo funciona:
 *  1. Inscripción (Google Forms): genera un código personal EXP-XXXX y lo envía por correo.
 *  2. Zonas: cada zona temática esconde un QR que abre la página del juego con ?z=1, ?z=2 o ?z=3.
 *     La zona es un capítulo con 3 enigmas (la 3 tiene 2 enigmas y la pregunta final).
 *  3. Cada enigma resuelto entrega una PALABRA de la frase final (8 palabras en total).
 *     El 3º enigma de los capítulos 1 y 2 revela además dónde está la siguiente zona.
 *  4. La pregunta final (enigma 9) pide rellenar los 8 huecos de la frase. Al acertar se cierra
 *     el caso, se anota la fecha y se envía un correo de confirmación.
 *  5. Sorteo: elige al azar entre quienes cerraron el caso.
 *
 * Las respuestas correctas viven SOLO en este script. La página las comprueba por POST (doPost).
 * Antes de usarlo: revisa VENTANA, ZONAS, PALABRAS, FRASE y PISTAS, y ejecuta configurar().
 */

// ID de la hoja de cálculo (déjalo vacío si el script está vinculado a la hoja)
const SS_ID = '17aOP6Qwm0fvryACe7rZw58nO20yAKN8ZuZdhBUSAsM0';

function ss_() {
  return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActive();
}

const CONFIG = {
  REMITENTE: 'Ego Sport Center',
  HOJA_CODIGOS: 'Codigos',
  HOJA_SORTEO: 'Sorteo',
  // Pestaña de respuestas del formulario de inscripción (renómbrala así en la hoja)
  HOJAS: { inscripcion: 'Inscripcion' },
  // Títulos EXACTOS de las preguntas del formulario de inscripción
  CAMPOS_INSCRIPCION: {
    nombre: 'Nombre y apellidos',
    socio: 'Número de socio',
    correo: 'Correo electrónico'
  },
  ENLACE_INSCRIPCION: 'PEGAR_ENLACE_INSCRIPCION',
  TEXTO_SORTEO: 'Te avisaremos por correo y en las redes del centro cuando se haga el sorteo.'
};

// Ventana de juego. null = sin límite. Madrid el 29/10/2026 es UTC+01:00 (el horario de verano acaba el 25/10).
// Ejemplo: { apertura: '2026-10-29T17:00:00+01:00', cierre: '2026-10-29T22:00:00+01:00' }
const VENTANA = { apertura: null, cierre: null };

// Nombre de cada zona/capítulo (los temas definitivos los decide el centro)
const ZONAS = { 1: 'El apagón', 2: 'Las sombras', 3: 'El culpable' };

// Las 8 palabras de la frase final, en el orden del enigma que las entrega (enigmas 1 a 8).
const PALABRAS = ['NOCHE', 'SOMBRA', 'ESPEJO', 'LIBRO', 'TUMBA', 'LUZ', 'VELA', 'MALDICIÓN'];

// Frase final con 8 huecos {1}..{8}. El hueco {n} se rellena con PALABRAS[n-1].
const FRASE = 'En la {1} de Halloween, la {2} salió del {3}, se llevó el {4} y lo enterró bajo una {5}. Solo la {6} de una {7} puede romper su {8}.';

const ORDEN = ['1.1', '1.2', '1.3', '2.1', '2.2', '2.3', '3.1', '3.2', 'FINAL'];

// Enigmas de cada capítulo, en orden. El 9º (FINAL) no está aquí: se responde con comprobarFrase.
const CAP_PISTAS = { 1: ['1.1', '1.2', '1.3'], 2: ['2.1', '2.2', '2.3'], 3: ['3.1', '3.2'] };

const MAX_INTENTOS = 10;  // fallos permitidos por enigma y persona
const AYUDA_CADA = 3;     // cada 3 fallos se desbloquea la siguiente ayuda (también se puede pedir)

/*
 * Cada enigma: cap, n (palabra de la frase que entrega, 1..8), tipo (cómo lo dibuja la página),
 * titulo, texto (las líneas que empiezan por "## " se muestran grandes, como mensaje cifrado),
 * datos (configuración del puzle interactivo), modo ('espejo' lo muestra invertido), ayudas (de la más suave a la más clara),
 * revela (solo en el último enigma de los capítulos 1 y 2: texto de la siguiente zona),
 * acepta (otras respuestas válidas además de la palabra).
 * Los marcados PROVISIONAL sirven para probar el sistema y se sustituyen uno a uno.
 */
const PISTAS = {
  '1.1': { cap: 1, n: 1, tipo: 'texto', titulo: 'El testimonio',
    texto: 'TESTIMONIO DE LA ÚLTIMA PERSONA QUE SALIÓ, copiado tal cual. Lee con atención cómo empieza cada línea.\n\nNadie oyó entrar a nadie, pero la puerta se abrió sola.\nOscuridad total durante tres minutos exactos.\nCuando volvió la luz, algo había desaparecido.\nHuellas mojadas en un suelo que estaba seco.\nEntre susurros, una voz repitió que aquello no había hecho más que empezar.',
    ayudas: ['Mira solo la primera letra de cada línea.', 'Une las cinco primeras letras, de arriba abajo.'] },
  '1.2': { cap: 1, n: 2, tipo: 'cesar', titulo: 'La declaración del camarero',
    texto: 'DECLARACIÓN DEL CAMARERO DE NOCHE\n\nAquella noche tenía seis copas servidas en la barra. En el apagón se rompió una.\n\nLo oí susurrar un nombre, siempre el mismo. Lo apunté tal como sonaba, pero antes moví cada letra hacia delante tantos lugares como copas quedaron enteras:\n\n## XTRGWF',
    ayudas: ['Antes de tocar las letras, averigua cuántas copas quedaron enteras: ese es el número de lugares.', 'El camarero movió las letras hacia delante, así que tú tienes que ir hacia atrás. Usa la rueda con un desplazamiento negativo.', 'Son 5 lugares hacia atrás: la X vuelve a ser la S. Sigue con las demás letras.'] },
  '1.3': { cap: 1, n: 3, tipo: 'sopa', titulo: 'La sopa del vigilante',
    texto: 'El vigilante nocturno dejó esta hoja con letras revueltas. Dice que en la lista hay ocho palabras de miedo, pero que hay una más, escondida y sin lista, que es la que importa.\n\nArrastra el dedo sobre las letras para marcar las palabras. Cuando encuentres la que no está en la lista, escríbela abajo.',
    datos: { lista: ["CALABAZA", "BRUJA", "VAMPIRO", "FANTASMA", "ZOMBI", "MOMIA", "CUERVO", "ESQUELETO"], rejilla: ["BITABRUJARSO", "AYLNROAHEOOJ", "COMAJOROUIHY", "RIAEOUASHVSP", "FVPHAOCJIATO", "SSGTVAMPIROR", "EEUFANTASMAJ", "OZOMBISNNUAE", "HEPQERHZONGM", "SOVREUCAIMOM", "JAAAZABALACU", "ESESQUELETOR"] },
    ayudas: ['Marca primero las ocho palabras de la lista. Las letras que queden sin marcar guardan el secreto.', 'La palabra escondida tiene 6 letras y va en diagonal, subiendo hacia la derecha.', 'La palabra empieza por la E del borde izquierdo de la fila 7 (contando desde arriba) y sube en diagonal hacia la derecha.'],
    revela: 'ZONA 2 · Las sombras\n\n[Pista del lugar donde está la zona 2: por definir con el centro]' },
  '2.1': { cap: 2, n: 4, tipo: 'deslizante', titulo: 'La portada rota',
    texto: 'Lo que desapareció aquella noche ha quedado hecho pedazos. Recompón la imagen deslizando las piezas hacia el hueco. Cuando esté completa, sabrás de qué se trata.',
    datos: { palabra: 'LIBRO' },
    ayudas: ['Empieza por los bordes: coloca la fila de arriba y luego ve bajando.', 'Solo se mueven las piezas que tocan el hueco. Ve despacio: mover sin plan suele deshacer lo ya hecho.', 'La respuesta es la palabra que verás escrita en la imagen terminada.'] },
  '2.2': { cap: 2, n: 5, tipo: 'texto', titulo: 'Donde descansa',
    texto: 'ENIGMA PROVISIONAL (prueba física de la zona), solo para probar el sistema. Aquí irá un puzle hecho a mano en la zona.\n\nEscribe la palabra: TUMBA',
    ayudas: ['Ayuda 1 de prueba: son 5 letras.', 'Ayuda 2 de prueba: empieza por T.'] },
  '2.3': { cap: 2, n: 6, tipo: 'deduccion', titulo: 'Los cinco testigos',
    texto: 'Cinco personas entraron en la sala, una tras otra, antes del apagón: Lucía, Mateo, Ulises, Nora y Zaira. Nadie recuerda el orden, pero sí estos detalles:\n\n• Entre Mateo y Nora llegó exactamente una persona.\n• Nora llegó justo antes que Zaira.\n• Lucía llegó antes que Ulises.\n• Nora no fue la tercera en llegar.\n• Ulises no llegó ni el primero ni el último.\n\nSolo vieron algo los que entraron en posición impar (1.º, 3.º y 5.º). Sus iniciales, en orden, forman la palabra que buscas.',
    datos: { nombres: ['Lucía', 'Mateo', 'Ulises', 'Nora', 'Zaira'] },
    ayudas: ['Empieza por Nora y Zaira: llegaron seguidas. ¿En qué posiciones pueden estar las dos?', 'Ulises está en la 2.ª, 3.ª o 4.ª posición, y Lucía llega antes que él.', 'El orden es Lucía, Mateo, Ulises, Nora y Zaira. Fíjate solo en las posiciones 1, 3 y 5.'],
    revela: 'ZONA 3 · El culpable\n\n[Pista del lugar donde está la zona 3: por definir con el centro]' },
  '3.1': { cap: 3, n: 7, tipo: 'candado', titulo: 'El candado de cuatro ruedas',
    texto: 'La caja donde se guardaba lo robado tenía un candado de cuatro ruedas. Cada rueda marca un número del 1 al 26, y cada número es una letra del abecedario (1 = A, 2 = B, 3 = C...).\n\nLas cuatro pistas del candado:\n\n1. Los tres minutos del apagón, multiplicados por siete, más uno.\n2. Las copas que quedaron enteras en la barra del camarero.\n3. Las campanadas de la medianoche.\n4. El número de culpables.\n\nGira las ruedas y lee la palabra que se forma.',
    datos: { ruedas: 4 },
    ayudas: ['La segunda pista habla del enigma del camarero: si lo resolviste, ya tienes ese número.', 'Los números son 22, 5, 12 y 1.', 'Pon 22, 5, 12 y 1 en las ruedas y lee las letras en orden.'] },
  '3.2': { cap: 3, n: 8, tipo: 'texto', titulo: 'El final del hechizo',
    texto: 'ENIGMA PROVISIONAL (prueba física de la zona), solo para probar el sistema. Aquí irá un puzle hecho a mano en la zona.\n\nEscribe la palabra: MALDICIÓN',
    ayudas: ['Ayuda 1 de prueba: son 9 letras.', 'Ayuda 2 de prueba: empieza por M.'] }
};

// Columnas de la hoja Codigos (empiezan en 1)
const COL = { CODIGO: 1, NOMBRE: 2, SOCIO: 3, CORREO: 4, ALTA: 5, CAP1: 6, CAP2: 7, CAP3: 8, ALERTAS: 9, PISTAS: 10, INTENTOS: 11 };
const ALFABETO_CODIGO = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';

/* ---------- Utilidades ---------- */

function normalizar_(texto) {
  return String(texto === undefined || texto === null ? '' : texto)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function normalizarCorreo_(texto) {
  return String(texto || '').trim().toLowerCase();
}

function valor_(valores, titulo) {
  const v = valores[titulo];
  return v && v.length ? String(v[0]).trim() : '';
}

function generarCodigo_(existentes) {
  for (let intento = 0; intento < 1000; intento++) {
    let sufijo = '';
    for (let i = 0; i < 4; i++) {
      sufijo += ALFABETO_CODIGO.charAt(Math.floor(Math.random() * ALFABETO_CODIGO.length));
    }
    const codigo = 'EXP-' + sufijo;
    if (!existentes[normalizar_(codigo)]) return codigo;
  }
  throw new Error('No se pudo generar un código único');
}

function enviarCorreo_(destino, asunto, cuerpo) {
  MailApp.sendEmail({ to: destino, subject: asunto, body: cuerpo, name: CONFIG.REMITENTE });
}

function hojaCodigos_() {
  const hoja = ss_().getSheetByName(CONFIG.HOJA_CODIGOS);
  if (!hoja) throw new Error('Falta la hoja ' + CONFIG.HOJA_CODIGOS + '. Ejecuta configurar().');
  return hoja;
}

/* ---------- Disparador del formulario de inscripción ---------- */

function alEnviarFormulario(e) {
  const hoja = e.range.getSheet().getName();
  if (hoja === CONFIG.HOJAS.inscripcion) return procesarInscripcion_(e.namedValues);
}

function procesarInscripcion_(valores) {
  const c = CONFIG.CAMPOS_INSCRIPCION;
  const nombre = valor_(valores, c.nombre);
  const socio = valor_(valores, c.socio);
  const correo = valor_(valores, c.correo);
  if (!correo) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const hoja = hojaCodigos_();
    const filas = hoja.getDataRange().getValues();
    const existentes = {};
    let codigo = null;
    for (let i = 1; i < filas.length; i++) {
      existentes[normalizar_(filas[i][COL.CODIGO - 1])] = true;
      if (normalizarCorreo_(filas[i][COL.CORREO - 1]) === normalizarCorreo_(correo)) {
        codigo = filas[i][COL.CODIGO - 1];
      }
    }
    if (!codigo) {
      codigo = generarCodigo_(existentes);
      hoja.appendRow([codigo, nombre, socio, correo, new Date(), '', '', '', '', '', '']);
    }
    enviarCorreo_(correo, 'El Expediente · Tu código de investigador',
      'Hola' + (nombre ? ' ' + nombre : '') + ',\n\n' +
      'Ya estás dentro del expediente. Tu código personal de investigador es:\n\n' +
      '    ' + codigo + '\n\n' +
      'Guárdalo y no lo compartas: lo necesitarás junto con este correo electrónico.\n\n' +
      'El juego es el jueves 29 de octubre en Ego Sport Center. Ven disfrazado, lleva tu móvil con internet ' +
      'y busca los códigos QR escondidos en las zonas del terror. Solo mayores de edad.\n\n' +
      '— ' + CONFIG.REMITENTE);
  } finally {
    lock.releaseLock();
  }
}

function anotarAlerta_(hoja, filaHoja, registro, texto) {
  const previas = registro[COL.ALERTAS - 1] ? registro[COL.ALERTAS - 1] + ' | ' : '';
  const nuevo = previas + texto;
  hoja.getRange(filaHoja, COL.ALERTAS).setValue(nuevo);
  registro[COL.ALERTAS - 1] = nuevo;
}

/* ---------- Sorteo ---------- */

function elegibles_() {
  const filas = hojaCodigos_().getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < filas.length; i++) {
    const f = filas[i];
    const completo = f[COL.CAP1 - 1] && f[COL.CAP2 - 1] && f[COL.CAP3 - 1];
    const bloqueado = String(f[COL.ALERTAS - 1] || '').indexOf('BLOQUEADO') !== -1;
    if (completo && !bloqueado) lista.push({ codigo: f[COL.CODIGO - 1], nombre: f[COL.NOMBRE - 1], socio: f[COL.SOCIO - 1], correo: f[COL.CORREO - 1] });
  }
  return lista;
}

function realizarSorteo() {
  const lista = elegibles_();
  if (!lista.length) throw new Error('No hay participantes elegibles todavía.');
  for (let i = lista.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = lista[i]; lista[i] = lista[j]; lista[j] = t;
  }
  const libro = ss_();
  let hoja = libro.getSheetByName(CONFIG.HOJA_SORTEO);
  if (!hoja) hoja = libro.insertSheet(CONFIG.HOJA_SORTEO);
  hoja.appendRow(['Sorteo realizado el', new Date(), 'Elegibles', lista.length]);
  const puestos = ['Ganador/a', 'Suplente 1', 'Suplente 2'];
  const resultado = [];
  for (let k = 0; k < puestos.length && k < lista.length; k++) {
    hoja.appendRow([puestos[k], lista[k].codigo, lista[k].nombre, lista[k].socio, lista[k].correo]);
    resultado.push(puestos[k] + ': ' + lista[k].nombre + ' (' + lista[k].codigo + ')');
  }
  return resultado;
}

/* ---------- Puesta en marcha ---------- */

function configurar() {
  const libro = ss_();
  let hoja = libro.getSheetByName(CONFIG.HOJA_CODIGOS);
  if (!hoja) {
    hoja = libro.insertSheet(CONFIG.HOJA_CODIGOS);
    hoja.appendRow(['Código', 'Nombre', 'Nº socio', 'Correo', 'Alta', 'Cap 1', 'Cap 2', 'Cap 3', 'Alertas', 'Pistas resueltas', 'Intentos']);
  }
  hoja.getRange(1, COL.PISTAS).setValue('Pistas resueltas');
  hoja.getRange(1, COL.INTENTOS).setValue('Intentos');
  const faltan = [];
  Object.keys(CONFIG.HOJAS).forEach(function (k) {
    if (!libro.getSheetByName(CONFIG.HOJAS[k])) faltan.push(CONFIG.HOJAS[k]);
  });
  const pendientes = [];
  Object.keys(PISTAS).forEach(function (id) {
    if (PISTAS[id].texto.indexOf('PROVISIONAL') !== -1) pendientes.push('Enigma ' + id + ' provisional');
    if (PISTAS[id].revela && PISTAS[id].revela.indexOf('[') !== -1) pendientes.push('Pista de zona del enigma ' + id + ' por definir');
  });
  if (PALABRAS.length !== 8) pendientes.push('La frase necesita exactamente 8 palabras');
  const existe = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'alEnviarFormulario';
  });
  if (!existe) {
    ScriptApp.newTrigger('alEnviarFormulario').forSpreadsheet(libro).onFormSubmit().create();
  }
  return { hojasQueFaltan: faltan, pendientesEnConfig: pendientes, disparadorCreado: !existe };
}

/* ---------- Montaje automático del formulario de inscripción (ejecutar UNA vez) ---------- */

function crearFormularios() {
  const libro = ss_();
  const c = CONFIG.CAMPOS_INSCRIPCION;
  const antes = libro.getSheets().map(function (h) { return h.getSheetId(); });
  const form = FormApp.create('El Expediente · Inscripción');
  form.setDescription('Solo mayores de edad. Recibirás tu código personal de investigador por correo (revisa spam).');
  form.setConfirmationMessage('Respuesta recibida. Te llegará un correo con tu código (revisa spam).');
  [c.nombre, c.socio, c.correo].forEach(function (t) { form.addTextItem().setTitle(t).setRequired(true); });
  ['Confirmo que soy mayor de edad', 'Acepto las bases y el tratamiento de mis datos para el juego y el sorteo'].forEach(function (t) {
    form.addCheckboxItem().setTitle(t).setChoiceValues(['Sí']).setRequired(true);
  });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, libro.getId());
  SpreadsheetApp.flush();
  const nueva = libro.getSheets().filter(function (h) { return antes.indexOf(h.getSheetId()) === -1; })[0];
  if (nueva) nueva.setName(CONFIG.HOJAS.inscripcion);
  const resumen = 'inscripcion -> ' + form.getPublishedUrl() + ' (editar: ' + form.getEditUrl() + ')';
  Logger.log(resumen);
  return resumen;
}

/* ====================================================================
 * JUEGO · API que consulta la página (GitHub Pages) por POST
 * Implementar > Nueva implementación > Aplicación web
 *   Ejecutar como: yo · Quién tiene acceso: cualquier persona
 * El QR de cada zona es:  <dirección de la página>?z=1  (?z=2, ?z=3)
 * ==================================================================== */

function leerResueltas_(registro) {
  return String(registro[COL.PISTAS - 1] || '').split(',').filter(function (x) { return x; });
}

function leerIntentos_(registro) {
  try { return JSON.parse(registro[COL.INTENTOS - 1] || '{}') || {}; } catch (err) { return {}; }
}

function guardar_(a, col, valor) {
  a.hoja.getRange(a.filaHoja, col).setValue(valor);
  a.registro[col - 1] = valor;
}

function respuestasValidas_(id) {
  const p = PISTAS[id];
  return [PALABRAS[p.n - 1]].concat(p.acepta || []).map(normalizar_);
}

function ventana_() {
  const ahora = new Date();
  if (VENTANA.apertura && ahora < new Date(VENTANA.apertura)) return 'El juego todavía no ha empezado. Vuelve el jueves 29 de octubre.';
  if (VENTANA.cierre && ahora > new Date(VENTANA.cierre)) return 'El juego ya ha terminado. ¡Gracias por participar!';
  return '';
}

// Comprueba código, correo, ventana de juego y zona anterior. Devuelve {fallo} o los datos de la fila.
function acceso_(codigo, correo, cap) {
  const NO = { ok: false, msg: 'Código o correo no válidos. Revisa el correo con tu código de investigador.' };
  if (!CAP_PISTAS[cap]) return { fallo: { ok: false, msg: 'Esta zona no existe.' } };
  if (!codigo || !correo) return { fallo: NO };

  const hoja = hojaCodigos_();
  const filas = hoja.getDataRange().getValues();
  let fila = -1;
  for (let i = 1; i < filas.length; i++) {
    if (normalizar_(filas[i][COL.CODIGO - 1]) === normalizar_(codigo)) { fila = i; break; }
  }
  if (fila === -1) return { fallo: NO };
  const registro = filas[fila];
  if (String(registro[COL.ALERTAS - 1] || '').indexOf('BLOQUEADO') !== -1) return { fallo: NO };
  if (normalizarCorreo_(registro[COL.CORREO - 1]) !== normalizarCorreo_(correo)) {
    anotarAlerta_(hoja, fila + 1, registro, 'Zona ' + cap + ': código usado con otro correo (' + correo + ')');
    return { fallo: NO };
  }
  const cerrada = ventana_();
  if (cerrada) return { fallo: { ok: false, msg: cerrada } };
  if (cap > 1 && !registro[COL.CAP1 - 1 + (cap - 2)]) {
    return { fallo: { ok: false, msg: 'Aún no has completado la zona anterior. Sigue las pistas hasta encontrar esta.' } };
  }
  return { hoja: hoja, filaHoja: fila + 1, registro: registro };
}

function vistaEnigma_(id, intentos) {
  const p = PISTAS[id];
  const fallos = intentos[id] || 0;
  const pedidas = intentos['h:' + id] || 0;
  const n = Math.min(p.ayudas.length, Math.max(Math.floor(fallos / AYUDA_CADA), pedidas));
  return {
    id: id, num: ORDEN.indexOf(id) + 1, de: ORDEN.length, tipo: p.tipo, titulo: p.titulo, texto: p.texto,
    modo: p.modo || 'normal', datos: p.datos || null, ayudas: p.ayudas.slice(0, n), masAyudas: n < p.ayudas.length, quedan: MAX_INTENTOS - fallos
  };
}

function fraseCompleta_() {
  return FRASE.replace(/\{(\d)\}/g, function (m, i) { return PALABRAS[Number(i) - 1]; });
}

// Estado que ve la persona en una zona: palabras conseguidas, enigma actual, pista de la siguiente zona o pregunta final.
function estado_(registro, cap) {
  const resueltas = leerResueltas_(registro), intentos = leerIntentos_(registro);
  const lista = CAP_PISTAS[cap];
  const hechas = lista.filter(function (p) { return resueltas.indexOf(p) !== -1; }).length;
  const capCompleto = hechas === lista.length;
  const juegoCompleto = !!registro[COL.CAP3 - 1];

  const palabras = [];
  Object.keys(PISTAS).forEach(function (id) {
    if (resueltas.indexOf(id) !== -1) palabras.push({ n: PISTAS[id].n, palabra: PALABRAS[PISTAS[id].n - 1] });
  });
  palabras.sort(function (x, y) { return x.n - y.n; });

  const pendiente = lista.filter(function (p) { return resueltas.indexOf(p) === -1; })[0];
  const ultimo = lista[lista.length - 1];
  const est = {
    cap: cap, zona: ZONAS[cap], total: lista.length, hechas: hechas, palabras: palabras,
    actual: pendiente ? vistaEnigma_(pendiente, intentos) : null,
    capCompleto: capCompleto, juegoCompleto: juegoCompleto,
    revela: cap < 3 && capCompleto ? (PISTAS[ultimo].revela || null) : null,
    frase: null, fraseCompleta: null, quedanFinal: MAX_INTENTOS - (intentos.F || 0)
  };
  if (cap === 3 && capCompleto && !juegoCompleto) {
    est.frase = {
      num: ORDEN.length, de: ORDEN.length, plantilla: FRASE,
      banco: palabras.map(function (x) { return x.palabra; }).sort(function (x, y) { return x.localeCompare(y, 'es'); })
    };
  }
  if (juegoCompleto) est.fraseCompleta = fraseCompleta_();
  return est;
}

function abrirZona(codigo, correo, cap) {
  cap = Number(cap);
  const a = acceso_(codigo, correo, cap);
  if (a.fallo) return a.fallo;
  return { ok: true, estado: estado_(a.registro, cap) };
}

function comprobarRespuesta(codigo, correo, id, respuesta) {
  id = String(id);
  if (!PISTAS[id]) return { ok: false, msg: 'Esta pista no existe.' };
  const cap = PISTAS[id].cap;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const a = acceso_(codigo, correo, cap);
    if (a.fallo) return a.fallo;
    const resueltas = leerResueltas_(a.registro);
    let palabra = null;
    if (resueltas.indexOf(id) === -1) {
      const pendiente = CAP_PISTAS[cap].filter(function (p) { return resueltas.indexOf(p) === -1; })[0];
      if (pendiente !== id) return { ok: false, msg: 'Resuelve antes el enigma anterior.' };
      const intentos = leerIntentos_(a.registro);
      const usados = intentos[id] || 0;
      if (usados >= MAX_INTENTOS) {
        return { ok: false, msg: 'Has agotado los intentos de este enigma. Habla con el personal del centro.' };
      }
      const r = normalizar_(respuesta);
      if (r === '' || respuestasValidas_(id).indexOf(r) === -1) {
        intentos[id] = usados + 1;
        guardar_(a, COL.INTENTOS, JSON.stringify(intentos));
        if (intentos[id] >= MAX_INTENTOS) anotarAlerta_(a.hoja, a.filaHoja, a.registro, 'Pista ' + id + ': ' + MAX_INTENTOS + ' intentos fallidos');
        return { ok: true, correcta: false, quedan: MAX_INTENTOS - intentos[id], estado: estado_(a.registro, cap) };
      }
      resueltas.push(id);
      guardar_(a, COL.PISTAS, resueltas.join(','));
      const lista = CAP_PISTAS[cap];
      if (cap < 3 && lista[lista.length - 1] === id) guardar_(a, COL.CAP1 + cap - 1, new Date());
      palabra = { n: PISTAS[id].n, palabra: PALABRAS[PISTAS[id].n - 1] };
    }
    return { ok: true, correcta: true, palabra: palabra, estado: estado_(a.registro, cap) };
  } finally {
    lock.releaseLock();
  }
}

// Desbloquea la siguiente ayuda del enigma actual sin necesidad de fallar.
function pedirAyuda(codigo, correo, id) {
  id = String(id);
  if (!PISTAS[id]) return { ok: false, msg: 'Esta pista no existe.' };
  const cap = PISTAS[id].cap;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const a = acceso_(codigo, correo, cap);
    if (a.fallo) return a.fallo;
    const resueltas = leerResueltas_(a.registro);
    const pendiente = CAP_PISTAS[cap].filter(function (p) { return resueltas.indexOf(p) === -1; })[0];
    if (pendiente === id) {
      const intentos = leerIntentos_(a.registro);
      const vista = vistaEnigma_(id, intentos);
      if (vista.masAyudas) {
        intentos['h:' + id] = vista.ayudas.length + 1;
        guardar_(a, COL.INTENTOS, JSON.stringify(intentos));
      }
    }
    return { ok: true, estado: estado_(a.registro, cap) };
  } finally {
    lock.releaseLock();
  }
}

// Enigma 9: rellenar los 8 huecos de la frase. 'lista' son las 8 palabras en el orden de los huecos.
function comprobarFrase(codigo, correo, lista) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const a = acceso_(codigo, correo, 3);
    if (a.fallo) return a.fallo;
    const resueltas = leerResueltas_(a.registro);
    if (!CAP_PISTAS[3].every(function (p) { return resueltas.indexOf(p) !== -1; })) {
      return { ok: false, msg: 'Resuelve antes los otros enigmas.' };
    }
    if (a.registro[COL.CAP3 - 1]) return { ok: true, correcta: true, estado: estado_(a.registro, 3) };
    if (!Array.isArray(lista) || lista.length !== PALABRAS.length) {
      return { ok: false, msg: 'Rellena los 8 huecos de la frase.' };
    }
    const intentos = leerIntentos_(a.registro);
    const usados = intentos.F || 0;
    if (usados >= MAX_INTENTOS) {
      return { ok: false, msg: 'Has agotado los intentos. Habla con el personal del centro.' };
    }
    const buena = lista.every(function (p, i) { return normalizar_(p) === normalizar_(PALABRAS[i]); });
    if (!buena) {
      intentos.F = usados + 1;
      guardar_(a, COL.INTENTOS, JSON.stringify(intentos));
      if (intentos.F >= MAX_INTENTOS) anotarAlerta_(a.hoja, a.filaHoja, a.registro, 'Frase final: ' + MAX_INTENTOS + ' intentos fallidos');
      return { ok: true, correcta: false, quedan: MAX_INTENTOS - intentos.F, estado: estado_(a.registro, 3) };
    }
    guardar_(a, COL.CAP3, new Date());
    try {
      enviarCorreo_(String(a.registro[COL.CORREO - 1]), 'El Expediente · Caso cerrado',
        '¡Caso cerrado, investigador!\n\n' + fraseCompleta_() + '\n\n' +
        'Has resuelto El Expediente y ya participas en el sorteo de la cena para dos.\n' + CONFIG.TEXTO_SORTEO + '\n\n— ' + CONFIG.REMITENTE);
    } catch (err) { /* si se agota la cuota de correo, el caso queda cerrado igualmente */ }
    return { ok: true, correcta: true, estado: estado_(a.registro, 3) };
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('El Expediente · API en marcha').setMimeType(ContentService.MimeType.TEXT);
}

// API JSON para la página alojada fuera de Google. Se consulta por POST (text/plain) y sin sesión de Google.
function doPost(e) {
  let res;
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.a === 'zona') res = abrirZona(d.c, d.m, d.z);
    else if (d.a === 'resp') res = comprobarRespuesta(d.c, d.m, d.id, d.v);
    else if (d.a === 'ayuda') res = pedirAyuda(d.c, d.m, d.id);
    else if (d.a === 'frase') res = comprobarFrase(d.c, d.m, d.v);
    else res = { ok: false, msg: 'Acción no válida.' };
  } catch (err) {
    res = { ok: false, msg: 'No se pudo comprobar. Inténtalo de nuevo.' };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}
