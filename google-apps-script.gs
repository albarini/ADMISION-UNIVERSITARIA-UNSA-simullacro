/**
 * SIMULACRO UNSA — Almacenamiento y análisis en Google Sheets
 * ============================================================
 * CÓMO INSTALARLO (5 minutos):
 *  1. Abre el editor de Apps Script (desde la hoja: Extensiones → Apps
 *     Script, o directamente en script.google.com — ambas formas sirven,
 *     el script ya está conectado a la hoja por su ID).
 *  2. (nada más que abrirlo)
 *  3. Borra todo el código anterior y pega ESTE archivo completo.
 *  4. Guarda (Ctrl+S).
 *  5. Implementar → Nueva implementación → tipo "Aplicación web":
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier persona
 *  6. Copia la URL que termina en /exec y pégala en la variable
 *     APPS_SCRIPT_URL del index.html.
 *  IMPORTANTE: cada vez que cambies este código debes crear una
 *  NUEVA implementación (o actualizar la versión), si no, los
 *  cambios no se aplican.
 *
 * QUÉ CREA AUTOMÁTICAMENTE:
 *  - Hoja "Resultados":   una fila por examen, coloreada por nivel.
 *  - Hoja "DetalleTemas": una fila por tema y por examen (para análisis).
 *  - Hoja "Resumen":      estadísticas automáticas — promedio general,
 *    ranking de temas más débiles, rendimiento por área y por carrera.
 */

/* ID de tu hoja de cálculo (sacado de su URL). Gracias a esto el script
   funciona aunque el proyecto de Apps Script no esté vinculado a la hoja. */
var SPREADSHEET_ID = '1woitAAesn5TJr6pp2f1MNobx2r3ACo-Jl60kyp25iac';

var HEADERS_RESULTADOS = ['Fecha','Nombre','Apellido','Área','Carrera','Curso','Puntaje','Nivel','Correctas','Incorrectas','Sin tiempo','Total','Seg/pregunta','Mejor racha','Mejor tema','Peor tema','Dispositivo'];
var HEADERS_DETALLE    = ['Fecha','Nombre','Apellido','Área','Carrera','Tema','Correctas','Preguntas','Precisión'];

/* Recibe los resultados del index.html nuevo (POST con JSON) */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    guardarResultado(data);
    return respuestaJSON({ ok: true });
  } catch (err) {
    return respuestaJSON({ ok: false, error: String(err) });
  }
}

/* Compatibilidad con la versión antigua (GET con parámetros) y prueba rápida */
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.ping) return respuestaJSON({ ok: true, ping: 'pong' });
    if (!p.nombre) return respuestaJSON({ ok: false, error: 'sin datos' });
    guardarResultado({
      fechaLocal: p.fecha || new Date().toLocaleString('es-PE'),
      nombre: p.nombre, apellido: p.apellido || '', area: p.area || '',
      carrera: p.carrera || '', curso: p.curso || 'Física',
      puntaje: Number(p.puntaje) || 0, nivel: p.nivel || '',
      correctas: Number(p.correctas) || 0, incorrectas: Number(p.incorrectas) || 0,
      sinTiempo: Number(p.sinTiempo) || 0, total: Number(p.total) || 0,
      colorFila: p.colorFila || '', temas: []
    });
    return respuestaJSON({ ok: true });
  } catch (err) {
    return respuestaJSON({ ok: false, error: String(err) });
  }
}

function respuestaJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function guardarResultado(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita filas mezcladas si dos alumnos terminan a la vez
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ---- Hoja Resultados ----
    var hoja = obtenerHoja(ss, 'Resultados', HEADERS_RESULTADOS);
    var fila = [
      data.fechaLocal || new Date().toLocaleString('es-PE'),
      data.nombre || '', data.apellido || '', data.area || '', data.carrera || '',
      data.curso || 'Física',
      Number(data.puntaje) || 0, data.nivel || '',
      Number(data.correctas) || 0, Number(data.incorrectas) || 0,
      Number(data.sinTiempo) || 0, Number(data.total) || 0,
      Number(data.tiempoMedio) || '', Number(data.mejorRacha) || '',
      data.mejorTema || '', data.peorTema || '', data.dispositivo || ''
    ];
    hoja.appendRow(fila);
    var r = hoja.getLastRow();
    if (data.colorFila) {
      hoja.getRange(r, 1, 1, HEADERS_RESULTADOS.length)
          .setBackground(data.colorFila).setFontColor('#1a1a1a');
    }
    // Ordenar por puntaje de mayor a menor (columna 7)
    if (hoja.getLastRow() > 2) {
      hoja.getRange(2, 1, hoja.getLastRow() - 1, HEADERS_RESULTADOS.length)
          .sort({ column: 7, ascending: false });
    }

    // ---- Hoja DetalleTemas (una fila por tema, para análisis fino) ----
    if (data.temas && data.temas.length) {
      var det = obtenerHoja(ss, 'DetalleTemas', HEADERS_DETALLE);
      var filas = data.temas.map(function (t) {
        return [
          data.fechaLocal || '', data.nombre || '', data.apellido || '',
          data.area || '', data.carrera || '',
          t.tema, Number(t.ok) || 0, Number(t.tot) || 0,
          t.tot ? Math.round((t.ok / t.tot) * 100) / 100 : 0
        ];
      });
      det.getRange(det.getLastRow() + 1, 1, filas.length, HEADERS_DETALLE.length).setValues(filas);
    }

    // ---- Hoja Resumen (se crea una sola vez, con fórmulas vivas) ----
    crearResumenSiFalta(ss);
  } finally {
    lock.releaseLock();
  }
}

function obtenerHoja(ss, nombre, headers) {
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#13131B').setFontColor('#DDB85C');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function crearResumenSiFalta(ss) {
  if (ss.getSheetByName('Resumen')) return;
  var h = ss.insertSheet('Resumen');
  h.getRange('A1').setValue('📊 RESUMEN GENERAL — SIMULACRO UNSA').setFontWeight('bold').setFontSize(14);

  h.getRange('A3').setValue('Total de exámenes rendidos:');
  h.getRange('B3').setFormula('=COUNTA(Resultados!B2:B)');
  h.getRange('A4').setValue('Puntaje promedio:');
  h.getRange('B4').setFormula('=IFERROR(ROUND(AVERAGE(Resultados!G2:G),1),0)');
  h.getRange('A5').setValue('Puntaje más alto:');
  h.getRange('B5').setFormula('=IFERROR(MAX(Resultados!G2:G),0)');
  h.getRange('A6').setValue('Estudiantes únicos (aprox.):');
  h.getRange('B6').setFormula('=IFERROR(COUNTA(UNIQUE(FILTER(Resultados!B2:B&" "&Resultados!C2:C,Resultados!B2:B<>""))),0)');

  h.getRange('A8').setValue('🔻 TEMAS MÁS DÉBILES (donde más hay que reforzar)').setFontWeight('bold');
  h.getRange('A9').setFormula('=IFERROR(QUERY(DetalleTemas!A2:I,"select F, sum(G), sum(H), round(sum(G)/sum(H)*100,1) where F is not null group by F order by sum(G)/sum(H) asc label F \'Tema\', sum(G) \'Correctas\', sum(H) \'Preguntas\', round(sum(G)/sum(H)*100,1) \'% acierto\'",0),"Aún no hay datos")');

  h.getRange('F8').setValue('🏛 RENDIMIENTO POR ÁREA').setFontWeight('bold');
  h.getRange('F9').setFormula('=IFERROR(QUERY(Resultados!A2:Q,"select D, count(D), round(avg(G),1) where D is not null group by D label D \'Área\', count(D) \'Exámenes\', round(avg(G),1) \'Promedio\'",0),"Aún no hay datos")');

  h.getRange('F15').setValue('🎓 RENDIMIENTO POR CARRERA').setFontWeight('bold');
  h.getRange('F16').setFormula('=IFERROR(QUERY(Resultados!A2:Q,"select E, count(E), round(avg(G),1) where E is not null group by E order by count(E) desc label E \'Carrera\', count(E) \'Exámenes\', round(avg(G),1) \'Promedio\'",0),"Aún no hay datos")');

  h.getRange('A1:K1').merge();
  h.setColumnWidths(1, 11, 140);
}
